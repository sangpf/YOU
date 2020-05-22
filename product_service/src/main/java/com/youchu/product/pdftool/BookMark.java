package com.youchu.product.pdftool;

import java.util.List;
import java.util.Stack;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 *
 * Word 文件中标签的封装类，保存了其定义和内部的操作
 */
public class BookMark {

    //以下为定义的常量

    /** 替换标签时，设于标签的后面   **/
    public static final int INSERT_AFTER = 0;

    /** 替换标签时，设于标签的前面   **/
    public static final int INSERT_BEFORE = 1;

    /** 替换标签时，将内容替换书签   **/
    public static final int REPLACE = 2;

    /** docx中定义的部分常量引用  **/
    public static final String RUN_NODE_NAME = "w:r";
    public static final String TEXT_NODE_NAME = "w:t";
    public static final String BOOKMARK_START_TAG = "bookmarkStart";
    public static final String BOOKMARK_END_TAG = "bookmarkEnd";
    public static final String BOOKMARK_ID_ATTR_NAME = "w:id";
    public static final String STYLE_NODE_NAME = "w:rPr";

    /** 内部的标签定义类  **/
    private CTBookmark _ctBookmark = null;

    /** 标签所处的段落  **/
    private XWPFParagraph _para = null;

    /** 标签所在的表cell对象  **/
    private XWPFTableCell _tableCell = null;

    /** 标签名称 **/
    private String _bookmarkName = null;

    /** 该标签是否处于表格内  **/
    private boolean _isCell = false;

    /**
     * 构造函数
     * @param ctBookmark
     * @param para
     */
    public BookMark(CTBookmark ctBookmark, XWPFParagraph para) {
        this._ctBookmark = ctBookmark;
        this._para = para;
        this._bookmarkName = ctBookmark.getName();
        this._tableCell = null;
        this._isCell = false;
    }

    /**
     * 构造函数，用于表格中的标签
     * @param ctBookmark
     * @param para
     * @param tableCell
     */
    public BookMark(CTBookmark ctBookmark, XWPFParagraph para, XWPFTableCell tableCell) {
        this(ctBookmark, para);
        this._tableCell = tableCell;
        this._isCell = true;
    }

    public boolean isInTable() {
        return this._isCell;
    }

    public XWPFTable getContainerTable() {
        return this._tableCell.getTableRow().getTable();
    }

    public XWPFTableRow getContainerTableRow() {
        return this._tableCell.getTableRow();
    }

    public String getBookmarkName() {
        return  this._bookmarkName;
    }

    public void insertTextAtBookMark(String bookmarkValue, int where) {
        //根据标签的类型，进行不同的操作
        if(this._isCell) {
            this.handleBookmarkedCells(bookmarkValue, where);
        } else {
            //普通标签，直接创建一个元素
            XWPFRun run = this._para.createRun();
            run.setText(bookmarkValue);
            switch(where) {
                case BookMark.INSERT_AFTER:
                    this.insertAfterBookmark(run);
                    break;
                case BookMark.INSERT_BEFORE:
                    this.insertBeforeBookmark(run);
                    break;
                case BookMark.REPLACE:
                    this.replaceBookmark(run);
                    break;
            }
        }
    }

    private void insertAfterBookmark(XWPFRun run) {
        Node nextNode = null;
        Node insertBeforeNode = null;
        Node styleNode = null;
        int bookmarkStartID = 0;
        int bookmarkEndID = -1;

        bookmarkStartID = this._ctBookmark.getId().intValue();
        nextNode = this._ctBookmark.getDomNode();
        while (bookmarkStartID != bookmarkEndID) {
            nextNode = nextNode.getNextSibling();
            if (nextNode.getNodeName().contains(BookMark.BOOKMARK_END_TAG)) {
                try {
                    bookmarkEndID = Integer.parseInt(
                            nextNode.getAttributes().getNamedItem(
                                    BookMark.BOOKMARK_ID_ATTR_NAME).getNodeValue());
                } catch (NumberFormatException nfe) {
                    bookmarkEndID = bookmarkStartID;
                }
            }
            else {
                if (nextNode.getNodeName().equals(BookMark.RUN_NODE_NAME)) {
                    styleNode = this.getStyleNode(nextNode);
                }
            }
        }
        insertBeforeNode = nextNode.getNextSibling();

        if (styleNode != null) {
            run.getCTR().getDomNode().insertBefore(
                    styleNode.cloneNode(true), run.getCTR().getDomNode().getFirstChild());
        }

        if (insertBeforeNode != null) {
            this._para.getCTP().getDomNode().insertBefore(
                    run.getCTR().getDomNode(), insertBeforeNode);
        }
    }

    private void insertBeforeBookmark(XWPFRun run) {
        Node insertBeforeNode = null;
        Node childNode = null;
        Node styleNode = null;

        insertBeforeNode = this._ctBookmark.getDomNode();
        childNode = insertBeforeNode.getPreviousSibling();

        // If a node is found, try to get the styling from it.
        if (childNode != null) {
            styleNode = this.getStyleNode(childNode);

            if (styleNode != null) {
                run.getCTR().getDomNode().insertBefore(
                        styleNode.cloneNode(true), run.getCTR().getDomNode().getFirstChild());
            }
        }

        this._para.getCTP().getDomNode().insertBefore(
                run.getCTR().getDomNode(), insertBeforeNode);
    }

    private void replaceBookmark(XWPFRun run) {
        Node nextNode = null;
        Node styleNode = null;
        Node lastRunNode = null;
        Node toDelete = null;
        NodeList childNodes = null;
        Stack<Node> nodeStack = null;
        boolean textNodeFound = false;
        boolean foundNested = true;
        int bookmarkStartID = 0;
        int bookmarkEndID = -1;
        int numChildNodes = 0;

        nodeStack = new Stack<Node>();
        bookmarkStartID = this._ctBookmark.getId().intValue();
        nextNode = this._ctBookmark.getDomNode();
        nodeStack.push(nextNode);

        // Loop through the nodes looking for a matching bookmarkEnd tag
        while (bookmarkStartID != bookmarkEndID) {
            nextNode = nextNode.getNextSibling();
            nodeStack.push(nextNode);

            // If an end tag is found, does it match the start tag? If so, end
            // the while loop.
            if (nextNode.getNodeName().contains(BookMark.BOOKMARK_END_TAG)) {
                try {
                    bookmarkEndID = Integer.parseInt(
                            nextNode.getAttributes().getNamedItem(
                                    BookMark.BOOKMARK_ID_ATTR_NAME).getNodeValue());
                } catch (NumberFormatException nfe) {
                    bookmarkEndID = bookmarkStartID;
                }
            }
            //else {
            // Place a reference to the node on the nodeStack
            //    nodeStack.push(nextNode);
            //}
        }

        if (!nodeStack.isEmpty()) {
            lastRunNode = nodeStack.peek();
            if ((lastRunNode.getNodeName().equals(BookMark.RUN_NODE_NAME))) {
                styleNode = this.getStyleNode(lastRunNode);
                if (styleNode != null) {
                    run.getCTR().getDomNode().insertBefore(
                            styleNode.cloneNode(true), run.getCTR().getDomNode().getFirstChild());
                }
            }

            this.deleteChildNodes(nodeStack);
        }

        // Place the text into position, between the bookmark tags.
        this._para.getCTP().getDomNode().insertBefore(
                run.getCTR().getDomNode(), nextNode);
    }

    private void deleteChildNodes(Stack<Node> nodeStack) {
        Node toDelete = null;
        int bookmarkStartID = 0;
        int bookmarkEndID = 0;
        boolean inNestedBookmark = false;

        // The first element in the list will be a bookmarkStart tag and that
        // must not be deleted.
        for(int i = 1; i < nodeStack.size(); i++) {
            // Get an element. If it is another bookmarkStart tag then
            // again, we do not want to delete it, it's matching end tag
            // or any nodes that fall inbetween.
            toDelete = nodeStack.elementAt(i);
            if(toDelete.getNodeName().contains(BookMark.BOOKMARK_START_TAG)) {
                bookmarkStartID = Integer.parseInt(
                        toDelete.getAttributes().getNamedItem(BookMark.BOOKMARK_ID_ATTR_NAME).getNodeValue());
                inNestedBookmark = true;
            }
            else if(toDelete.getNodeName().contains(BookMark.BOOKMARK_END_TAG)) {
                bookmarkEndID = Integer.parseInt(
                        toDelete.getAttributes().getNamedItem(BookMark.BOOKMARK_ID_ATTR_NAME).getNodeValue());
                if(bookmarkEndID == bookmarkStartID) {
                    inNestedBookmark = false;
                }
            }
            else {
                if(!inNestedBookmark) {
                    this._para.getCTP().getDomNode().removeChild(toDelete);
                }
            }
        }
    }

    private Node getStyleNode(Node parentNode) {
        Node childNode = null;
        Node styleNode = null;
        if (parentNode != null) {
            if (parentNode.getNodeName().equalsIgnoreCase(BookMark.RUN_NODE_NAME)
                    && parentNode.hasChildNodes()) {
                childNode = parentNode.getFirstChild();
                if (childNode.getNodeName().equals("w:rPr")) {
                    styleNode = childNode;
                } else {
                    while ((childNode = childNode.getNextSibling()) != null) {
                        if (childNode.getNodeName().equals(BookMark.STYLE_NODE_NAME)) {
                            styleNode = childNode;
                            // Note setting to null here if a style node is
                            // found in order order to terminate any further
                            // checking
                            childNode = null;
                        }
                    }
                }
            }
        }
        return (styleNode);
    }

    public String getBookmarkText() throws XmlException {
        StringBuilder builder = null;
        // Are we dealing with a bookmarked table cell? If so, the entire
        // contents of the cell - if anything - must be recovered and returned.
        if(this._tableCell != null) {
            builder = new StringBuilder(this._tableCell.getText());
        }
        else {
            builder = this.getTextFromBookmark();
        }
        return(builder == null ? null : builder.toString());
    }

    private StringBuilder getTextFromBookmark() throws XmlException {
        int startBookmarkID = 0;
        int endBookmarkID = -1;
        Node nextNode = null;
        Node childNode = null;
        CTText text = null;
        StringBuilder builder = null;
        String rawXML = null;

        // Get the ID of the bookmark from it's start tag, the DOM node from the
        // bookmark (to make looping easier) and initialise the StringBuilder.
        startBookmarkID = this._ctBookmark.getId().intValue();
        nextNode = this._ctBookmark.getDomNode();
        builder = new StringBuilder();

        // Loop through the nodes held between the bookmark's start and end
        // tags.
        while (startBookmarkID != endBookmarkID) {

            // Get the next node and, if it is a bookmarkEnd tag, get it's ID
            // as matching ids will terminate the while loop..
            nextNode = nextNode.getNextSibling();
            if (nextNode.getNodeName().contains(BookMark.BOOKMARK_END_TAG)) {
                try {
                    endBookmarkID = Integer.parseInt(
                            nextNode.getAttributes().
                                    getNamedItem(BookMark.BOOKMARK_ID_ATTR_NAME).getNodeValue());
                } catch (NumberFormatException nfe) {
                    endBookmarkID = startBookmarkID;
                }
            } else {
                if (nextNode.getNodeName().equals(BookMark.RUN_NODE_NAME)
                        && nextNode.hasChildNodes()) {
                    // Get the text from the child nodes.
                    builder.append(this.getTextFromChildNodes(nextNode));
                }
            }
        }
        return (builder);
    }

    private String getTextFromChildNodes(Node node) throws XmlException {
        NodeList childNodes = null;
        Node childNode = null;
        CTText text = null;
        StringBuilder builder = new StringBuilder();
        int numChildNodes = 0;

        // Get a list of chid nodes from the node passed to the method and
        // find out how many children there are in the list.
        childNodes = node.getChildNodes();
        numChildNodes = childNodes.getLength();

        // Iterate through the children one at a time - it is possible for a
        // run to ciontain zero, one or more text nodes - and recover the text
        // from an text type child nodes.
        for (int i = 0; i < numChildNodes; i++) {
            // Get a node and check it's name. If this is 'w:t' then process as
            // text type node.
            childNode = childNodes.item(i);
            if (childNode.getNodeName().equals(BookMark.TEXT_NODE_NAME)) {
                // If the node reports it's type as txet, then simply call the
                // getNodeValue() method to get at it's text.
                if (childNode.getNodeType() == Node.TEXT_NODE) {
                    builder.append(childNode.getNodeValue());
                } else {
                    // Correct the type by parsing the node's XML markup and
                    // creating a CTText object. Call the getStringValue()
                    // method on that to get the text.
                    text = CTText.Factory.parse(childNode);
                    builder.append(text.getStringValue());
                }
            }
        }
        return (builder.toString());
    }
    private void handleBookmarkedCells(String bookmarkValue, int where) {
        List<XWPFParagraph> paraList = null;
        List<XWPFRun> runs = null;
        XWPFParagraph para = null;
        XWPFRun readRun = null;
        // Get a list if paragraphs from the table cell and remove any and all.
        paraList = this._tableCell.getParagraphs();
        for(int i = 0; i < paraList.size(); i++) {
            this._tableCell.removeParagraph(i);
        }
        para = this._tableCell.addParagraph();
        para.createRun().setText(bookmarkValue);
    }
}
