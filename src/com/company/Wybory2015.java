package com.company;

        import java.io.IOException;
        import org.jsoup.Jsoup;
        import org.jsoup.nodes.Document;
        import org.jsoup.nodes.Element;
        import org.jsoup.select.Elements;
        import java.io.File;
        import java.io.FileInputStream;
        import java.io.FileNotFoundException;
        import java.io.FileOutputStream;
        import java.io.InputStream;
        import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
        import org.apache.poi.ss.usermodel.Cell;
        import org.apache.poi.ss.usermodel.CellType;
        import org.apache.poi.ss.usermodel.Sheet;
        import org.apache.poi.ss.usermodel.Workbook;
        import org.apache.poi.ss.usermodel.WorkbookFactory;
        import java.util.*;
/**
 *
 * @author PR
 */
class HtmlTableToArray {
    public List<List<String>> rowToArray(Elements tableRowElements) {
        List<List<String>> output = new ArrayList<>();
        for (Element row : tableRowElements) {
            List<String> columns = new ArrayList<>();
            Elements rowItems = row.select("td:not(:has(*))");
            if (rowItems.size() > 1) {
                for (Element rowItem : rowItems) {
                    columns.add(rowItem.text());
                }
                output.add(columns);
            }
        }
    return output;
    }
}

class DetailedTable {
    public List<List<String>> getDetails(String html){
        List<List<String>> output = new ArrayList<>();
        try {
            Document doc = Jsoup.connect(html).get();
            Element tableElements = doc.select("table").get(0);
            Elements tableRowElements = tableElements.select(":not(thead) tr");
            HtmlTableToArray tableToArray = new HtmlTableToArray();
            output=tableToArray.rowToArray(tableRowElements);
        }
        catch (IOException e) {
            e.printStackTrace();
        }
        return output;
    }
}

class ListsTable {
    public List<List<String>> getLists(String html){
        List<List<String>> output = new ArrayList<>();
        try {
            Document doc = Jsoup.connect(html).get();
            Element tableElements = doc.select("table.jstable").get(1);
            Elements tableRowElements = tableElements.select(":not(thead) tr");
            HtmlTableToArray tableToArray = new HtmlTableToArray();
            output=tableToArray.rowToArray(tableRowElements);
        }
        catch (IOException e) {
            e.printStackTrace();}
        return output;
    }
}

class Parser{
    public String kodTer(Document doc){
        String col = doc.select("div:containsOwn(Kod terytorialny)").get(0).nextElementSibling().text();
        return col;
    }
    public String wojew(Document doc){
        String col = doc.select("div:containsOwn(Województwo)").get(0).nextElementSibling().text();
        return col;
    }
    public String powiat(Document doc){
        String col = doc.select("div:containsOwn(Powiat)").get(0).nextElementSibling().text();
        return col;
    }
    public String gmina(Document doc){
        String col = doc.select("div:containsOwn(Gmina)").get(0).nextElementSibling().text();
        return col;
    }
    public String numerObw(Document doc){
        String col = doc.select("div:containsOwn(Numer obwodu)").get(0).nextElementSibling().text();
        return col;
    }
    public String adres(Document doc){
        String col = doc.select("div:containsOwn(Adres)").get(0).nextElementSibling().text();
        return col;
    }
    public String liczbaWyb(Document doc){
        String col = doc.select("div:containsOwn(Liczba wyborców uprawnionych do głosowania (umieszczonych w spisie, z uwzględnieniem dodatkowych formularzy) w chwili zakończenia głosowania)").get(0).nextElementSibling().text();
        return col;
    }
    public String otrzymKarty(Document doc){
        String col = doc.select("div:containsOwn(Komisja otrzymała kart do głosowania)").get(0).nextElementSibling().text();
        return col;
    }
    public String nieWykKarty(Document doc){
        String col = doc.select("div:containsOwn(Nie wykorzystano kart do głosowania)").get(0).nextElementSibling().text();
        return col;
    }
    public String wybKarty(Document doc){
        String col = doc.select("div:containsOwn(Liczba wyborców, którym wydano karty do głosowania (liczba podpisów w spisie oraz adnotacje \"odmowa podpisu\"))").get(0).nextElementSibling().text();
        return col;
    }
    public String wybPelnom(Document doc){
        String col = doc.select("div:containsOwn(Liczba wyborców głosujących przez pełnomocnika (liczba kart do głosowania wydanych na podstawie otrzymanych przez komisję aktów pełnomocnictwa))").get(0).nextElementSibling().text();
        return col;
    }
    public String wybZasw(Document doc){
        String col = doc.select("div:containsOwn(Liczba wyborców głosujących na podstawie zaświadczenia o prawie do głosowania)").get(0).nextElementSibling().text();
        return col;
    }
    public String wybPakiet(Document doc){
        String col = doc.select("div:containsOwn(Liczba wyborców, którym wysłano pakiety wyborcze)").get(0).nextElementSibling().text();
        return col;
    }
    public String kopZwr(Document doc){
        String col = doc.select("div:containsOwn(Liczba otrzymanych kopert zwrotnych)").get(0).nextElementSibling().text();
        return col;
    }
    public String kopZwrBezOsw(Document doc){
        String col = doc.select("div:containsOwn(Liczba kopert zwrotnych, w których nie było oświadczenia o osobistym i tajnym oddaniu głosu)").get(0).nextElementSibling().text();
        return col;
    }
    public String kopZwrOswBezPodp(Document doc){
        String col = doc.select("div:containsOwn(Liczba kopert zwrotnych, w których oświadczenie nie było podpisane przez wyborcę)").get(0).nextElementSibling().text();
        return col;
    }
    public String kopZwrBezKop(Document doc){
        String col = doc.select("div:containsOwn(Liczba kopert zwrotnych, w których nie było koperty na karty do głosowania)").get(0).nextElementSibling().text();
        return col;
    }
    public String kopZwrNiezakKop(Document doc){
        String col = doc.select("div:containsOwn(Liczba kopert zwrotnych, w których znajdowała się niezaklejona koperta na karty do głosowania)").get(0).nextElementSibling().text();
        return col;
    }
    public String kopDoUrny(Document doc){
        String col = doc.select("div:containsOwn(Liczba kopert na karty do głosowania wrzuconych do urny)").get(0).nextElementSibling().text();
        return col;
    }
    public String kartyZUrny(Document doc){
        String col = doc.select("div:containsOwn(Liczba kart wyjętych z urny)").get(0).nextElementSibling().text();
        return col;
    }
    public String kartyZKopert(Document doc){
        String col = doc.select("div:containsOwn(w tym liczba kart wyjętych z kopert na karty do głosowania)").get(0).nextElementSibling().text();
        return col;
    }
    public String kartyNiewazne(Document doc){
        String col = doc.select("div:containsOwn(Liczba kart nieważnych (innych niż urzędowo ustalone lub nieopatrzonych pieczęcią obwodowej komisji wyborczej))").get(0).nextElementSibling().text();
        return col;
    }
    public String kartyWazne(Document doc){
        String col = doc.select("div:containsOwn(Liczba kart ważnych)").get(0).nextElementSibling().text();
        return col;
    }
    public String glosyNiewazne(Document doc){
        String col = doc.select("div:containsOwn(Liczba głosów nieważnych (z kart ważnych))").get(0).nextElementSibling().text();
        return col;
    }
    public String xPRzy2(Document doc){
        String col = doc.select("div:containsOwn(w tym z powodu postawienia znaku \"X\" obok nazwiska dwóch lub większej liczby kandydatów z różnych list)").get(0).nextElementSibling().text();
        return col;
    }
    public String xPRzy0(Document doc){
        String col = doc.select("div:containsOwn(w tym z powodu niepostawienia znaku \"X\" obok nazwiska żadnego kandydata z którejkolwiek z list)").get(0).nextElementSibling().text();
        return col;
    }
    public String xPRzyUniew(Document doc){
        String col = doc.select("div:containsOwn(w tym z powodu postawienia znaku \"X\" wyłącznie obok nazwiska kandydata z listy, której rejestracja została unieważniona (art. 227 § 3 Kodeksu))").get(0).nextElementSibling().text();
        return col;
    }
    public String glosyWazne(Document doc){
        String col = doc.select("div:containsOwn(Liczba głosów ważnych oddanych łącznie na wszystkie listy kandydatów (z kart ważnych))").get(0).nextElementSibling().text();
        return col;
    }
}

class GeneralData {
    public List<String> getGenData (String html){
        List<String> output = new ArrayList<>();
        Parser parser = new Parser();
        try{
            Document doc = Jsoup.connect(html).get();

            output.add(parser.kodTer(doc));
            output.add(parser.wojew(doc));
            output.add(parser.powiat(doc));
            output.add(parser.gmina(doc));
            output.add(parser.numerObw(doc));
            output.add(parser.adres(doc));
            output.add(parser.liczbaWyb(doc));
            output.add(parser.otrzymKarty(doc));
            output.add(parser.nieWykKarty(doc));
            output.add(parser.wybKarty(doc));
            output.add(parser.wybPelnom(doc));
            output.add(parser.wybZasw(doc));
            output.add(parser.wybPakiet(doc));
            output.add(parser.kopZwr(doc));
            output.add(parser.kopZwrBezOsw(doc));
            output.add(parser.kopZwrOswBezPodp(doc));
            output.add(parser.kopZwrBezKop(doc));
            output.add(parser.kopZwrNiezakKop(doc));
            output.add(parser.kopDoUrny(doc));
            output.add(parser.kartyZUrny(doc));
            output.add(parser.kartyZKopert(doc));
            output.add(parser.kartyNiewazne(doc));
            output.add(parser.kartyWazne(doc));
            output.add(parser.glosyNiewazne(doc));
            output.add(parser.xPRzy2(doc));
            output.add(parser.xPRzy0(doc));
            output.add(parser.xPRzyUniew(doc));
            output.add(parser.glosyWazne(doc));
        }
        catch (IOException e) {
            e.printStackTrace();}
        catch (NullPointerException e) {
            System.out.print("Caught the NullPointerException");
            //zapisać w logu
        }
        return output;
    }
}

class ArrayToExcel {
    public int daneWgList(List<List<String>> listsData, List<String> genData, String html, Sheet sheet3, int row3Index) {
        for (List<String> row : listsData) {
            org.apache.poi.ss.usermodel.Row row3 = sheet3.getRow(row3Index++);
            if (row3 == null)
                row3 = sheet3.createRow(row3Index - 1);
            if (row.get(0) == null) {
                row3Index--;
                break;
            }
            int k = 0;
            for (String text : row) {
                for (int m = 0; m < 6; m++) {
                    Cell cell3 = row3.getCell(m);
                    if (cell3 == null)
                        cell3 = row3.createCell(m);
                    cell3.setCellType(CellType.STRING);
                    cell3.setCellValue(genData.get(m));
                }
                Cell cell3 = row3.getCell(6);
                if (cell3 == null)
                    cell3 = row3.createCell(6);
                cell3.setCellType(CellType.STRING);
                cell3.setCellValue(html);

                cell3 = row3.getCell(7 + k);
                if (cell3 == null)
                    cell3 = row3.createCell(7 + k);
                if (k == 2) {
                    cell3.setCellType(CellType.NUMERIC);
                } else {
                    cell3.setCellType(CellType.STRING);
                }
                cell3.setCellValue(text);
                k++;
            }
        }
        return row3Index;
    }

    public int daneOgolne(List<String> genData, String html, Sheet sheet1, int row1Index) {
        org.apache.poi.ss.usermodel.Row row1 = sheet1.getRow(row1Index++);
        if (row1 == null)
            row1 = sheet1.createRow(row1Index - 1);
        int m = 0;
        for (String data : genData) {
            Cell cell1 = row1.getCell(m);
            if (cell1 == null)
                cell1 = row1.createCell(m);
            cell1.setCellType(CellType.STRING);
            cell1.setCellValue(data);
            m++;
        }
        Cell cell1 = row1.getCell(m);
        if (cell1 == null)
            cell1 = row1.createCell(m);
        cell1.setCellType(CellType.STRING);
        cell1.setCellValue(html);
        return row1Index;
    }

    public int daneSzczeg(List<List<String>> data, String html, Sheet sheet2, int row2Index) {
        for (List<String> dataRow : data) {
            org.apache.poi.ss.usermodel.Row row2 = sheet2.getRow(row2Index++);
            if (row2 == null)
                row2 = sheet2.createRow(row2Index - 1);
            if (dataRow.get(0) == null) {
                row2Index--;
                break;
            }
            int k = 0;
            for (String rowElem : dataRow) {
                Cell cell2 = row2.getCell(6);
                if (cell2 == null)
                    cell2 = row2.createCell(6);
                cell2.setCellType(CellType.STRING);
                cell2.setCellValue(html);
                cell2 = row2.getCell(7 + k);
                if (cell2 == null)
                    cell2 = row2.createCell(7 + k);
                cell2.setCellType(CellType.STRING);
                cell2.setCellValue(rowElem);
                k++;
            }
        }
        return row2Index;
    }
}

public class Wybory2015 {

    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {
        File src = new File("c:\\DANE\\FIRMA\\INFORMATYKA\\JAVA tools\\Wybory 2015 IntelliJ.xlsx");
        InputStream inp = new FileInputStream(src);
        Workbook wb = WorkbookFactory.create(inp);

        ArrayToExcel atx = new ArrayToExcel();
        Scanner keyboard = new Scanner(System.in);
        System.out.println("Podaj liczbę obwodów (od 1 do 27859):");
        int cycles = keyboard.nextInt();

        Sheet sheet0 = wb.getSheetAt(0);
        Sheet sheet1 = wb.getSheetAt(1);
        Sheet sheet2 = wb.getSheetAt(2);
        Sheet sheet3 = wb.getSheetAt(3);

        DetailedTable dt = new DetailedTable();
        GeneralData gd = new GeneralData();
        ListsTable lt = new ListsTable();
        int row1Index=1;
        int row2Index=1;
        int row3Index=1;

        for (int i=0;i<cycles;i++){//27859
            long begin = System.nanoTime();
            org.apache.poi.ss.usermodel.Row row0 = sheet0.getRow(i);
            Cell cell0 = row0.getCell(0);
            String html = cell0.getStringCellValue();

            List<List<String>> listsData = new ArrayList<>();
            listsData=lt.getLists(html);

            List<String> genData = new ArrayList<>();
            genData = gd.getGenData(html);

            List<List<String>> detData = new ArrayList<>();
            detData=dt.getDetails(html);

            row1Index = atx.daneOgolne(genData,html,sheet1,row1Index);
            row2Index = atx.daneSzczeg(detData, html, sheet2,row2Index);
            row3Index = atx.daneWgList(listsData,genData,html,sheet3,row3Index);

            FileOutputStream fileOut = new FileOutputStream(src);
            wb.write(fileOut);
            fileOut.close();
            long end = System.nanoTime();

            System.out.println(i);
            System.out.println(String.format("%.2f",(((float)end-(float)begin)/1000000000)) + "s");
        }
    }
}








