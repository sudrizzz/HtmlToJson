
public class Main {

    public static void main(String[] args) {
        String filepath = "layout.html";
        Cell cell = new Cell();
        System.out.println(cell.getTableInfo(filepath, "GBK"));
    }
}
