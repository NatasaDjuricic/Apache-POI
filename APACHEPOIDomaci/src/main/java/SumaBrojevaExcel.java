import java.io.IOException;

public class SumaBrojevaExcel {
    public static void main(String[] args) throws IOException {

        ExcelReader excelReader = new ExcelReader("C:\\Users\\natad\\IdeaProjects\\APACHEPOIDomaci\\TextBook.xlsx");


        int sum = 0;

        System.out.print("In 1st column of TextBook 'Brojevi' Sheet, numbers are: ");

        for (int i = 1; i <= excelReader.getLastRow(); i++) {
            int broj = excelReader.getIntegerData("Brojevi", i, 0);
            sum += i;

            if (i != excelReader.getLastRow()) {
                System.out.print(broj + ", ");
            }
            else {
                System.out.print(broj);
            }

        }
        System.out.println("\nSum of all nubers in 1st row is: " + sum);
    }
}
