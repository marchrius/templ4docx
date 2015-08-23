package pl.jsolve.templ4docx.variable;

import java.util.List;
import java.util.Map;

import pl.jsolve.sweetener.collection.Collections;
import pl.jsolve.sweetener.collection.Maps;
import pl.jsolve.templ4docx.util.Key;

public class Variables {

    private Map<String, TextVariable> textVariables;
    private Map<String, ImageVariable> imageVariables;
    private List<TableVariable> tableVariables;
    private Map<String, BulletListVariable> bulletListVariables;
    private Map<String, ConditionVariable> conditionVariables;

    public Variables() {
        this.textVariables = Maps.newHashMap();
        this.imageVariables = Maps.newHashMap();
        this.tableVariables = Collections.newArrayList();
        this.bulletListVariables = Maps.newHashMap();
        this.conditionVariables = Maps.newHashMap();
    }

    public TextVariable addTextVariable(TextVariable textVariable) {
        return this.textVariables.put(textVariable.getKey(), textVariable);
    }

    public ImageVariable addImageVariable(ImageVariable imageVariable) {
        return this.imageVariables.put(imageVariable.getKey(), imageVariable);
    }

    public TableVariable addTableVariable(TableVariable tableVariable) {
        this.tableVariables.add(tableVariable);
        return tableVariable;
    }

    public BulletListVariable addBulletListVariable(BulletListVariable bulletListVariable) {
        this.bulletListVariables.put(bulletListVariable.getKey(), bulletListVariable);
        return bulletListVariable;
    }

    public ConditionVariable addConditionVariable(ConditionVariable conditionVariable) {
        this.conditionVariables.put(conditionVariable.getKey(), conditionVariable);
        return conditionVariable;
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

    public Map<String, ConditionVariable> getConditionVariables() {
        return conditionVariables;
    }

    public ConditionVariable getCondition(String key) {
        return conditionVariables.get(key);
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
        }
        return null; // TODO: throw exception
    }

}
