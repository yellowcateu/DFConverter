/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package components;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author zolty
 */
public class ReadXLSFile {
    
    private File xlsFile;
    private String[] ColumnsID = {"IndeksLeku-","IloscSprzedana-","NazwaLeku--------------------------------------------------","CenaDetaliczna--","CenaDetBrutto-","Vat-","KwotaVat---"};
    private String LekID, LekOfeIlosc, LekNazwa, LekCenaDet, LekCenaDetBr, LekVat, LekKwotaVat;

    
    public void ReadXLSFile(){

    }
    
    public void initVars(){

    }
    
    public void setXLSFile(String path){
        this.xlsFile = new File(path);
    }
    
    public File getXLSFile(){
        return this.xlsFile;
    }

    public String StartReadig(File xlsfile) throws FileNotFoundException, IOException{
        String resultDF = "";
        String dataFarmLine = "-";
        try {

            FileInputStream file = new FileInputStream(xlsfile);

            //Get the workbook instance for XLS file 
            HSSFWorkbook workbook;
            workbook = new HSSFWorkbook(file);

            //Get first sheet from the workbook
            HSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();
            double tempdouble;
            
            while(rowIterator.hasNext()) {
                Row row = rowIterator.next();
                if(row.getRowNum()==0){
                    row = rowIterator.next();
                    for(int i=0; i<ColumnsID.length; i++){
                        dataFarmLine += ColumnsID[i];
                    }
                    dataFarmLine += "\r\n";
                }

                //For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while(cellIterator.hasNext()) {
                    
                //Methods for converting excell file to a data-farm format
                    
                    Cell cell = cellIterator.next();
                    //indeks leku do kazdej pozycji na razie pusty
                    if(cell.getColumnIndex()==0){
                        
                        LekID = buildSpaces(ColumnsID[0].length());
                        //System.out.println(LekID);
                        resultDF += LekID+"\t\t";
                        
                    }
                    //nazwa handlowa leku
                    if(cell.getColumnIndex()==5){
                        
                        if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
                            tempdouble = cell.getNumericCellValue();  
                            LekNazwa = doSpaces(ColumnsID[2].length(), String.format("%.3f", tempdouble).replace(',', '.'), false);
                            resultDF += tempdouble+ "\t\t";
                        }else if(cell.getCellType()==Cell.CELL_TYPE_STRING){
                            LekNazwa = doSpaces(ColumnsID[2].length()-1,cell.getStringCellValue(),true);
                            resultDF += cell.getStringCellValue() + "\t\t";
                        }
                        
                    }// oferowana ilosc
                    else if (cell.getColumnIndex() == 8){
                        if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
                            tempdouble = cell.getNumericCellValue();                            
                            LekOfeIlosc = doSpaces(ColumnsID[1].length(), String.format("%.2f", tempdouble).replace(',', '.'), false);
                            resultDF += tempdouble+ "\t\t";
                        }else if(cell.getCellType()==Cell.CELL_TYPE_STRING){
                            LekOfeIlosc = doSpaces(ColumnsID[1].length()-1,cell.getStringCellValue(),true);
                            resultDF += cell.getStringCellValue() + "\t\t";
                        }
                    }//cena detaliczna leku
                    else if (cell.getColumnIndex() == 10){
                        if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
                            tempdouble = cell.getNumericCellValue();  
                            LekCenaDet = doSpaces(ColumnsID[3].length(), String.format("%.2f", tempdouble).replace(',', '.'), false);
                            resultDF += tempdouble+ "\t\t";
                        }else if(cell.getCellType()==Cell.CELL_TYPE_STRING){
                            LekCenaDet = doSpaces(ColumnsID[3].length()-1,cell.getStringCellValue(),true);
                            resultDF += cell.getStringCellValue() + "\t\t";
                        }
                    }//vat
                    else if (cell.getColumnIndex() == 11){
                        if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
                            tempdouble = cell.getNumericCellValue();
                            tempdouble = setPrecision(tempdouble, 2)*100;  
                            int x;
                            x = Integer.parseInt(String.valueOf(tempdouble).substring(0, 1));
                            LekVat = doSpaces(ColumnsID[5].length(), String.valueOf(x)+"%", false);
                            resultDF += tempdouble+ "\t\t";
                        }else if(cell.getCellType()==Cell.CELL_TYPE_STRING){
                            LekVat = doSpaces(ColumnsID[5].length()-1,cell.getStringCellValue(),true);
                            resultDF += cell.getStringCellValue() + "\t\t";
                        }
                    }//cena detaliczna leku Brutto
                    else if (cell.getColumnIndex() == 12){
                        if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
                            tempdouble = cell.getNumericCellValue(); 
                            LekCenaDetBr = doSpaces(ColumnsID[4].length(), String.format("%.2f", tempdouble).replace(',', '.'), false);
                            resultDF += tempdouble+ "\t\t";
                        }else if(cell.getCellType()==Cell.CELL_TYPE_STRING){
                            LekCenaDetBr = doSpaces(ColumnsID[4].length()-1,cell.getStringCellValue(),true);
                            resultDF += cell.getStringCellValue() + "\t\t";
                        }
                    }//kwota Vat
                    else if (cell.getColumnIndex() == 15){
                        if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
                            tempdouble = cell.getNumericCellValue(); 
                            LekKwotaVat = doSpaces(ColumnsID[6].length(), String.format("%.2f", tempdouble).replace(',', '.'), false);
                            resultDF += tempdouble+ "\t\t";
                        }else if(cell.getCellType()==Cell.CELL_TYPE_STRING){
                            LekKwotaVat = doSpaces(ColumnsID[6].length()-1,cell.getStringCellValue(),true);
                            resultDF += cell.getStringCellValue() + "\t\t";
                        }
                    }

                }
                dataFarmLine += LekID+LekOfeIlosc+LekNazwa+LekCenaDet+LekCenaDetBr+LekVat+LekKwotaVat+"\r\n";
                resultDF += "\r\n";
            }
            file.close();
           /*
            FileOutputStream out = 
                new FileOutputStream(new File("C:\\test.xls"));
            workbook.write(out);
            out.close();
            */
        } catch (FileNotFoundException e) {
        } catch (IOException e) {
        }
        return dataFarmLine;
    }
    
    private String doSpaces(int colLength, String valFromExcell, boolean isString){
        String result, addedSpaces;
        result = "";
        addedSpaces = "";
        int spaces = 0;
        
        spaces = colLength - valFromExcell.length();
//        
//        if(isString == true){
//            addedSpaces = buildSpaces(spaces-1);
//        }else{
//            addedSpaces = buildSpaces(spaces);
//        }
        addedSpaces = buildSpaces(spaces);
        if(isString == true){
            result = " "+valFromExcell+addedSpaces;
        }else{
            result = addedSpaces+valFromExcell;
        }
                
        return result;
    }
    
    public static double setPrecision(double value, int precision) {
        BigDecimal A = new BigDecimal(value + "");
        return A.setScale(precision, BigDecimal.ROUND_HALF_UP).doubleValue();
    }
    
    public String buildSpaces(int len){
        StringBuilder sb = new StringBuilder();

        for(int i=0; i < len; i++) {
            sb.append(' ');
        }

        return sb.toString();
    }
    
}
