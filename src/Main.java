public class Main {

    public static void main(String[] args) {
        String filepath = "layout3.html";
        Cell cell = new Cell();
        System.out.println(cell.getTableInfo(filepath, "GBK").get("datajson"));
        System.out.println(cell.getTableInfo(filepath, "GBK").get("pluginjson"));
    }
}
