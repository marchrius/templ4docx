package pl.jsolve.templ4docx.variable;

import java.io.File;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import pl.jsolve.sweetener.collection.Collections;
import pl.jsolve.sweetener.collection.Maps;
import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.util.Key;

public class Variables {

  private Map<String, TextVariable> textVariables;
  private Map<String, ImageVariable> imageVariables;
  private List<TableVariable> tableVariables;
  private Map<String, BulletListVariable> bulletListVariables;
  private Map<String, ObjectVariable> objectVariables;
  private Map<String, DocumentVariable> documentVariables;

  public Variables() {
    this.textVariables = Maps.newHashMap();
    this.imageVariables = Maps.newHashMap();
    this.tableVariables = Collections.newArrayList();
    this.bulletListVariables = Maps.newHashMap();
    this.objectVariables = Maps.newHashMap();
    this.documentVariables = Maps.newHashMap();
  }

  public TextVariable addTextVariable(TextVariable textVariable) {
    return this.textVariables.put(textVariable.getKey(), textVariable);
  }

  public TextVariable addTextVariable(String key, String value) {
    TextVariable textVariable = new TextVariable(key, value);
    return add(textVariable);
  }

  public ImageVariable addImageVariable(ImageVariable imageVariable) {
    return this.imageVariables.put(imageVariable.getKey(), imageVariable);
  }

  public ImageVariable addImageVariable(String key, String imagePath, int width, int height) {
    ImageVariable imageVariable = new ImageVariable(key, imagePath, width, height);
    return add(imageVariable);
  }

  public ImageVariable addImageVariable(String key, File imageFile, int width, int height) {
    ImageVariable imageVariable = new ImageVariable(key, imageFile, width, height);
    return add(imageVariable);
  }

  public ImageVariable addImageVariable(String key, String imagePath, ImageType imageType, int width, int height) {
    ImageVariable imageVariable = new ImageVariable(key, imagePath, imageType, width, height);
    return add(imageVariable);
  }

  public ImageVariable addImageVariable(String key, File imageFile, ImageType imageType, int width, int height) {
    ImageVariable imageVariable = new ImageVariable(key, imageFile, imageType, width, height);
    return add(imageVariable);
  }

  public TableVariable addTableVariable(TableVariable tableVariable) {
    this.tableVariables.add(tableVariable);
    return tableVariable;
  }

  public BulletListVariable addBulletListVariable(BulletListVariable bulletListVariable) {
    this.bulletListVariables.put(bulletListVariable.getKey(), bulletListVariable);
    return bulletListVariable;
  }

  public BulletListVariable addBulletListVariable(String key, List<? extends Variable> variables) {
    BulletListVariable bulletListVariable = new BulletListVariable(key, variables);
    return add(bulletListVariable);
  }

  public DocumentVariable addDocumentVariable(DocumentVariable documentVariable) {
    return this.documentVariables.put(documentVariable.getKey(), documentVariable);
  }

  public DocumentVariable addDocumentVariable(String key, Docx document) {
    return addDocumentVariable(key, document.getXWPFDocument());
  }

  public DocumentVariable addDocumentVariable(String key, XWPFDocument document) {
    return addDocumentVariable(key, document, true);
  }

  public DocumentVariable addDocumentVariable(String key, Docx document, boolean asUniqueParagraph) {
    return addDocumentVariable(key, document.getXWPFDocument(), asUniqueParagraph);
  }

  public DocumentVariable addDocumentVariable(String key, XWPFDocument document, boolean asUniqueParagraph) {
    DocumentVariable documentVariable = new DocumentVariable(key, document, asUniqueParagraph);
    return add(documentVariable);
  }

  public List<ObjectVariable> addObjectVariable(ObjectVariable objectVariable) {
    List<ObjectVariable> tree = Collections.newArrayList();
    tree.add(objectVariable);
    tree.addAll(objectVariable.getFieldVariablesTree());
    for (ObjectVariable var : tree) {
      this.objectVariables.put(var.getKey(), var);
    }
    return tree;
  }

  public Map<String, TextVariable> getTextVariables() {
    return textVariables;
  }

  public Map<String, ImageVariable> getImageVariables() {
    return imageVariables;
  }

  public List<TableVariable> getTableVariables() {
    return tableVariables;
  }

  public Map<String, BulletListVariable> getBulletListVariables() {
    return bulletListVariables;
  }

  public Map<String, DocumentVariable> getDocumentVariables() {
    return documentVariables;
  }

  public Map<String, ObjectVariable> getObjectVariables() {
    return objectVariables;
  }

  public Variable getVariable(Key key) {
    switch (key.getVariableType()) {
      case TEXT:
        return textVariables.get(key.getKey());
      case IMAGE:
        return imageVariables.get(key.getKey());
      case TABLE:
        for (Key subkey : key.getSubKeys()) {
          for (TableVariable tableVariable : tableVariables) {
            if (tableVariable.containsKey(subkey.getKey())) {
              return tableVariable;
            }
          }
        }
        break;
      case BULLET_LIST:
        return bulletListVariables.get(key.getKey());
      case OBJECT:
        return objectVariables.get(key.getKey());
      case DOCUMENT:
        return documentVariables.get(key.getKey());
    }
    return null; // TODO: throw exception
  }



  @SuppressWarnings("unchecked")
  public <T extends Variable> T add(T variable) {
    if (variable instanceof TextVariable) {
      return (T) addTextVariable((TextVariable) variable);
    } else if (variable instanceof ImageVariable) {
      return (T) addImageVariable((ImageVariable) variable);
    } else if (variable instanceof TableVariable) {
      return (T) addTableVariable((TableVariable) variable);
    } else if (variable instanceof BulletListVariable) {
      return (T) addBulletListVariable((BulletListVariable) variable);
    } else if (variable instanceof ObjectVariable) {
      return (T) addObjectVariable((ObjectVariable) variable);
    } else if (variable instanceof DocumentVariable) {
      return (T) addDocumentVariable((DocumentVariable) variable);
    }
    return null; // TODO: throw exception
  }

  public Variables clone(boolean alsoBulletListInsert, boolean alsoDocumentInsert) {
    Variables newV = new Variables();
    for (String key : getTextVariables().keySet()) {
      newV.add(getTextVariables().get(key));
    }
    for (String key : getImageVariables().keySet()) {
      newV.add(getImageVariables().get(key));
    }
    if (alsoBulletListInsert) {
      for (String key : getBulletListVariables().keySet()) {
        newV.add(getBulletListVariables().get(key));
      }
    }
    for (String key : getObjectVariables().keySet()) {
      newV.add(getObjectVariables().get(key));
    }
    if (alsoDocumentInsert) {
      for (String key : getDocumentVariables().keySet()) {
        newV.add(getDocumentVariables().get(key));
      }
    }
    for (TableVariable tableVariable : getTableVariables()) {
      newV.add(tableVariable);
    }
    return newV;
  }
}
