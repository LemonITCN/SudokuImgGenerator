package cn.lemonit.play;

/**
 * 数独信息对象
 *
 * @author LemonIT.CN
 */
public class SudokuInfo {

    /**
     * 数独编号
     */
    private String number;
    /**
     * 数独题目内容
     */
    private String content;

    public SudokuInfo() {
    }

    public SudokuInfo(String number, String content) {
        this.number = number;
        this.content = content;
    }

    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }
}
