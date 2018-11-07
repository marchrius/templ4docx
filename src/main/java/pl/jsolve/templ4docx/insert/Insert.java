package pl.jsolve.templ4docx.insert;

import pl.jsolve.templ4docx.util.Key;

/**
 * Insert which is equal to one found variable.
 * @author Lukasz Stypka
 */
public abstract class Insert {

    /**
     * Key - information about found variable
     */
    protected Key key;

    /**
     * boolean - if the insert is created from BulletList
     */
    protected boolean inAList;

    public Insert(Key key) {
        this.key = key;
    }

    public Key getKey() {
        return key;
    }

    public void setKey(Key key) {
        this.key = key;
    }

    public boolean isInAList() {
        return inAList;
    }

    public void setInAList(boolean inAList) {
        this.inAList = inAList;
    }

    @Override
    public String toString() {
			return "Insert [key=" + key + ", inAList=" + inAList + "]";
    }

}
