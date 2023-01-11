package model;

import java.io.Serializable;

public class PageModel implements Serializable {
    private static final long serialVersionUID = 7160343888629944529L;
    private String pageTitle;
    private String pageText;
    private int pageNumber;

    public String getPageTitle() {
        return pageTitle;
    }

    public void setPageTitle(String pageTitle) {
        this.pageTitle = pageTitle;
    }

    public String getPageText() {
        return pageText;
    }

    public void setPageText(String pageText) {
        this.pageText = pageText;
    }

    public int getPageNumber() {
        return pageNumber;
    }

    public void setPageNumber(int pageNumber) {
        this.pageNumber = pageNumber;
    }
}
