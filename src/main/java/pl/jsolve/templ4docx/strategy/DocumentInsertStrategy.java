package pl.jsolve.templ4docx.strategy;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;

import pl.jsolve.templ4docx.insert.DocumentInsert;
import pl.jsolve.templ4docx.insert.Insert;
import pl.jsolve.templ4docx.variable.DocumentVariable;
import pl.jsolve.templ4docx.variable.Variable;

public class DocumentInsertStrategy implements InsertStrategy{

	@Override
	public void insert(Insert insert, Variable variable) {
		if(!(insert instanceof DocumentInsert)){
			return;
		}
		if(!(variable instanceof DocumentVariable)){
			return;
		}
		
		DocumentInsert documentInsert = (DocumentInsert) insert;
		DocumentVariable documentVariable = (DocumentVariable) variable;
		XWPFParagraph templateParagraph = documentInsert.getParagraph();
		List<IBodyElement> bodyElements = getReverseListOfBodyElements(documentVariable.getDocument());
		XmlCursor cursor = getCursorFromParagraph(templateParagraph);
		
		for(IBodyElement bodyElement : bodyElements){
			BodyElementType bodyElementType = bodyElement.getElementType();
			
			if(bodyElementType.name().equals("PARAGRAPH")){
				XWPFParagraph pr = (XWPFParagraph) bodyElement;
				XWPFParagraph newPr = templateParagraph.getBody().insertNewParagraph(cursor);
				cloneParagraph(newPr, pr);
                cursor = newPr.getCTP().newCursor();
				
			} else if(bodyElementType.name().equals("TABLE")){
				XWPFTable table = (XWPFTable) bodyElement;
                XWPFTable newTable = templateParagraph.getDocument().insertNewTbl(cursor);
                cloneTable(newTable, table);
                cursor = newTable.getCTTbl().newCursor();
			}
		}
		clean(templateParagraph, documentInsert);
	}
	
	private void cloneParagraph(XWPFParagraph clone, XWPFParagraph source) {
        CTPPr pPr = clone.getCTP().addNewPPr();
        pPr.set(source.getCTP().getPPr());
        for (XWPFRun r : source.getRuns()) {
            XWPFRun newRun = clone.createRun();
            cloneRun(newRun, r);
        }
    }

    private void cloneRun(XWPFRun clone, XWPFRun source) {
        CTRPr rPr = clone.getCTR().addNewRPr();
        rPr.set(source.getCTR().getRPr());
        clone.setText(source.getText(0));
    }
    
    private void cloneTable(XWPFTable clone, XWPFTable source) {
    	CTTblPr tblPr = clone.getCTTbl().addNewTblPr();
    	tblPr.set(source.getCTTbl().getTblPr());
    	
    	for(int i = 0; i < source.getRows().size(); i ++){
    	
    		XWPFTableRow newRow = clone.getRow(i);
    		
    		if(newRow == null){
    			newRow = clone.createRow();
    		}
    		
    		cloneRow(newRow, source.getRows().get(i));
    	}
    }
	
	private void cloneRow(XWPFTableRow clone, XWPFTableRow source) {
		CTRow ctRow = clone.getCtRow();
		ctRow.set(source.getCtRow());
	}

	private XmlCursor getCursorFromParagraph(XWPFParagraph paragraph) {
		return paragraph.getCTP().newCursor();
	}

	private List<IBodyElement> getReverseListOfBodyElements(XWPFDocument document){
		List<IBodyElement> bodyElements = new ArrayList<IBodyElement>(document.getBodyElements());
		Collections.reverse(bodyElements);
		return bodyElements;		
	}
	
	private void clean(XWPFParagraph paragraph, DocumentInsert insert){
		for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);
            if (StringUtils.contains(text, insert.getKey().getKey())) {
                text = StringUtils.replace(text, insert.getKey().getKey(), "");
                run.setText(text, 0);
            }
        }
	}
	
}
