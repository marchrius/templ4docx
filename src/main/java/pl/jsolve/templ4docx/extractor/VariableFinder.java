package pl.jsolve.templ4docx.extractor;

import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import pl.jsolve.sweetener.collection.Collections;
import pl.jsolve.sweetener.collection.Maps;
import pl.jsolve.templ4docx.cleaner.ParagraphCleaner;
import pl.jsolve.templ4docx.cleaner.TableRowCleaner;
import pl.jsolve.templ4docx.insert.BulletListInsert;
import pl.jsolve.templ4docx.insert.DocumentInsert;
import pl.jsolve.templ4docx.insert.ImageInsert;
import pl.jsolve.templ4docx.insert.Insert;
import pl.jsolve.templ4docx.insert.TableCellInsert;
import pl.jsolve.templ4docx.insert.TableRowInsert;
import pl.jsolve.templ4docx.insert.TextInsert;
import pl.jsolve.templ4docx.strategy.InsertStrategyChooser;
import pl.jsolve.templ4docx.util.Key;
import pl.jsolve.templ4docx.variable.Variables;

/**
 * Utility class responsible for preparing list of inserts. For each variable found in template file there is creating
 * appropriate insert (TextInsert, ImageInsert, BulletListInser or TableInsert)
 * @author Lukasz Stypka
 */
public class VariableFinder {

    private InsertStrategyChooser insertStrategyChooser;
    private TableRowCleaner tableRowCleaner;
    private ParagraphCleaner paragraphCleaner;
    private KeyExtractor keyExtractor;

    public VariableFinder(Variables variables) {
        this.tableRowCleaner = new TableRowCleaner();
        this.paragraphCleaner = new ParagraphCleaner();
        this.keyExtractor = new KeyExtractor();
        this.insertStrategyChooser = new InsertStrategyChooser(variables, tableRowCleaner, paragraphCleaner);
    }

    /**
     * Returns list of inserts for found variables.
     * @param document Apache POI XWPFDocument object
     * @param variables Object which contains list of variables grouped by type
     * @return List of inserts
     */
    public List<Insert> find(XWPFDocument document, Variables variables) {
        List<Insert> inserts = Collections.newArrayList();
        List<Key> keys = keyExtractor.extractKeys(variables);
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            inserts.addAll(find(paragraph, document, null, keys));
        }

        findInTables(inserts, document.getTables(), keys);

        mergeTableInserts(inserts, variables);
        return inserts;
    }

    /**
     * Finds variables recursively for each table in the template file
     * @param inserts List<Insert> inserts
     * @param tables List<XWPFTable> tables
     * @param keys List<Key> keys
     */
    private void findInTables(List<Insert> inserts, List<XWPFTable> tables, List<Key> keys) {
        for (XWPFTable tbl : tables) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    if (!cell.getTables().isEmpty()) {
                        findInTables(inserts, cell.getTables(), keys);
                    }
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        inserts.addAll(find(paragraph, null, cell, keys));
                    }
                }
            }
        }
    }

    /**
     * Finds variables in given paragraph
     * @param paragraph XWPFParagraph paragraph
     * @param document  XWPFDocument document
     * @param cell XWPFTableCell cell
     * @param keys  List<Key> keys
     * @return
     */
    private List<Insert> find(XWPFParagraph paragraph, XWPFDocument document, XWPFTableCell cell, List<Key> keys) {
        List<Insert> inserts = Collections.newArrayList();
        StringBuilder sb = new StringBuilder();
        for (XWPFRun run : paragraph.getRuns()) {
            sb.append(run.getText(0));
        }
        String paragraphText = sb.toString();

        for (Key key : keys) {
            if (StringUtils.contains(paragraphText, key.getKey())) {
                switch (key.getVariableType()) {
                case TEXT:
                    inserts.add(new TextInsert(key, paragraph));
                    break;
                case IMAGE:
                    inserts.add(new ImageInsert(key, paragraph));
                    break;
                case BULLET_LIST:
                    inserts.add(new BulletListInsert(key, paragraph, cell, document));
                    break;
                case TABLE:
                    if (cell != null) {
                        inserts.add(new TableCellInsert(key, cell));
                    }
                    break;
                case DOCUMENT:
                	inserts.add(new DocumentInsert(key, paragraph));
                	break;
                }
            }
        }
        return inserts;
    }

    /**
     * This method checks whether many Table Inserts belong to the same row. If so, cell inserts which belong to the
     * same row are transform to one TableRowInsert.
     * @param inserts List<Insert> inserts
     * @param variables Variables variables
     */
    private void mergeTableInserts(List<Insert> inserts, Variables variables) {
        Map<XWPFTableRow, TableRowInsert> rowInserts = Maps.newHashMap();
        List<Insert> insertsToRemove = Collections.newArrayList();
        for (Insert insert : inserts) {
            if (insert instanceof TableCellInsert) {
                TableCellInsert cellInsert = (TableCellInsert) insert;
                insertsToRemove.add(cellInsert);
                XWPFTableRow tableRow = cellInsert.getCell().getTableRow();
                // if rowInsert contains reference to XWPFTableRow, it'll mean that appropriate TableRowInsert already
                // exist. In this case add this cell Insert to found row Insert
                if (rowInserts.containsKey(tableRow)) {
                    TableRowInsert tableRowInsert = rowInserts.get(tableRow);
                    tableRowInsert.add(cellInsert);
                } else {
                    // otherwise, create new table row insert
                    TableRowInsert rowInsert = new TableRowInsert();
                    rowInsert.add(cellInsert);
                    rowInserts.put(tableRow, rowInsert);
                }

            }
        }

        // Remove old cell inserts from main list of inserts
        for (Insert insert : insertsToRemove) {
            inserts.remove(insert);
        }
        // and add list of row inserts
        inserts.addAll(rowInserts.values());
    }

    /**
     * Execute appropriate strategy for each insert. This method replace found variable (Insert) to appropriate strategy
     * (replacing text, insert image, insert new row, insert bullet list)
     * @param inserts Insert inserts
     */
    public void replace(List<Insert> inserts) {
        for (Insert insert : inserts) {
            insertStrategyChooser.replace(insert);
        }
        insertStrategyChooser.cleanUp();
    }

}