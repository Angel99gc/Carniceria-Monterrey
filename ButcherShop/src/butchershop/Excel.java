/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package butchershop;

/**
 *
 * @author Angel
 */

 
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import static java.lang.String.format;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.util.concurrent.TimeUnit;


public class Excel   {
        Workbook wordBook;
        String file;
        int hoja;
        String FechaHoy;
        public Excel(String fileName, String fecha)
        {  
                      

            try{
                            
                FileInputStream fileInputStream = new FileInputStream(fileName);
                this.FechaHoy = fecha;
                this.file = fileName;
                this.hoja = Integer.parseInt(fecha.split("-")[0])-1;
           
                Workbook wb = WorkbookFactory.create(fileInputStream);

                this.wordBook = wb;


                
            }catch(Exception e){
                System.out.println("Error: "+e.toString());
            }
          
            
        }
        /*
        *Obtiene Lista de precios segun Ventas/Gasto
        *@autor: Angel Gamboa Cruz
        *@param: boolean isVenta  -Si es true es una venta SINO un gasto
        *@return: ArrayList<String> Lista  -Lista con cada precio apuntado.
        */
        public ArrayList getPrecios(boolean isVenta){
            ArrayList Lista = new ArrayList();
            Sheet hssfSheet = this.wordBook.getSheetAt(this.hoja);
            int tipo = 1;
            if(isVenta)tipo=0;
            for(int i = 1; i<110 ; i++){
                Row row = (Row) hssfSheet.getRow(i);//fila
                Cell cell = (Cell) row.getCell(tipo); //columna
                
                if(cell.getCellType() == 0) {
                  
                    int valorNum = (int) cell.getNumericCellValue();
                    String valor = String.valueOf(valorNum);
                    Lista.add(valor);
                }
                else{
                    
                    break;
                }
            }
            System.out.println(Lista);
            return Lista;
         
        }
          /*
        *Inserta Lista de precios segun Ventas/Gasto
        *@autor: Angel Gamboa Cruz
        *@param: boolean isVenta  -Si es true es una venta SINO un gasto
        *@param: ArrayList<String> list  -Lista con cada precio apuntado.
        */
        public void setPrecios(boolean isVenta, ArrayList<String> list){
            Sheet hssfSheet = this.wordBook.getSheetAt(this.hoja);
            int tipo = 1;
            if(isVenta)tipo=0;
            
            for(int i = 1; i<110 ; i++){
                Row row = (Row) hssfSheet.getRow(i);//fila
                Cell cell = (Cell) row.getCell(tipo); //columna
                if(i-1==list.size()){
                    cell.setCellValue("");
                    break;
                };
                cell.setCellValue(Integer.parseInt(list.get(i-1)));

            }
         
        }
        /*
        *Actualiza el documento en excel. 
        *@autor: Angel Gamboa Cruz
        *@param:
        *@return:
        */
        public void saveWordBook(){
            try (FileOutputStream fileOut = new FileOutputStream(file)){
                this.wordBook.write(fileOut);
            } catch (Exception e) {
                System.out.println("Error: "+e.toString());
            }
        }
         /*
        *Cambiar una celda en la posicion segun tipoSet y lo cambia por dato
        *@autor: Angel Gamboa Cruz
        *@param: String tipoSet  -Tipo de celda a cambiar("Billetes","Monedas","Caja"),
        *@param: int dato -Dato a cambiar en la celda.
        *@return:
        */
        public void setData(String tipoSet, int dato){
            Cell xcell = this.getCellSpecific(tipoSet,this.getHoja(true));
            xcell.setCellValue(dato);
        }
         /*
        *Cambiar una celda en la posicion segun tipoSet y lo cambia por dato
        *@autor: Angel Gamboa Cruz
        *@param: String tipoSet  -Tipo de celda a cambiar("Billetes","Monedas","Caja"),
        *@return:
        */
        public void setCaja(MenuInicio menu, int dato){
            int indexHoja = this.getHoja(false);
            if(indexHoja==50){//nuevo mes, nuevo documento
                Excel sigMes = menu.newDocument(this.tomorrowDate());
                sigMes.setCaja(menu, dato);
                sigMes.saveWordBook();
            }else if(indexHoja== 51){//Es sabado, nuevo documento para Lunes
                Excel sigLunes = menu.newDocument(this.afterSundayDate());
                sigLunes.setCaja(menu, dato);
                sigLunes.saveWordBook();
            }else{
                Cell xcell = this.getCellSpecific("Caja",indexHoja);
                xcell.setCellValue(dato);
            }
        }
         /*
        *Obtener una celda en la posicion segun tipoSet.
        *@autor: Angel Gamboa Cruz
        *@param: String tipoSet  -Tipo de celda a cambiar("Billetes","Monedas","Caja"),
        *@return: int valor  -Valor en el campo espeficado por tipoSet
        */
        public int getData(String tipoSet){
            Cell xcell = this.getCellSpecific(tipoSet, this.getHoja(true));
            return (int) xcell.getNumericCellValue();
        }
        /*
        *Obtiene una unica celda del documento.
        *@autor: Angel Gamboa Cruz
        *@param: String tipoSet  -Tipo de celda a obtener("Billetes","Monedas","Caja"),
        *@param: Cell cell -Unica celda obtenida
        *@return:
        */
        public Cell getCellSpecific(String tipoSet, int hoja){
            int row=0;
            int cell=0;
            switch (tipoSet) {
                case "Billetes":
                    row=0;
                    cell=5;
                    break;
                case "Monedas":
                    row=1;
                    cell=5;
                    break;
                case "Tarjeta":
                    row=8;
                    cell=4;
                    break;
                case "Caja":
                    row=2;
                    cell=5;
                    break;
                case "TotalDerecho":
                    row=3;
                    cell=5;
                    break;
                case "TotalVentas":
                    row=1;
                    cell=3;
                    break;
                case "TotalGastos":
                    row=2;
                    cell=3;
                    break;
                case "TotalIzquierda":
                    row=3;
                    cell=3;
                    break;
                case "Diferencia":
                    row=5;
                    cell=4;
                    break;
                case "TotalDeposito":
                   row=10;
                   cell=4;
                   break;
                case "TotalEfectivo":
                   row=7;
                   cell=4;
                   break;
                case "TotalTarjeta":
                   row=8;
                   cell=4;
                   break;
                default:
                    break;
            }
            Sheet xSheet = this.wordBook.getSheetAt(hoja);
            Row xrow = (Row) xSheet.getRow(row);//fila
            Cell xcell = (Cell) xrow.getCell(cell); //columna
            return xcell;
        }
        public void updateFormulas(){
            FormulaEvaluator evaluator =  this.wordBook.getCreationHelper().createFormulaEvaluator();
            Cell c = getCellSpecific("TotalDerecho",this.getHoja(true));
            evaluator.evaluateFormulaCell(c);
            Cell c2 = getCellSpecific("TotalVentas",this.getHoja(true));
            evaluator.evaluateFormulaCell(c2);
            Cell c3 = getCellSpecific("TotalGastos",this.getHoja(true));
            evaluator.evaluateFormulaCell(c3);
            Cell c4 = getCellSpecific("TotalIzquierda",this.getHoja(true));
            evaluator.evaluateFormulaCell(c4);
            Cell c5 = getCellSpecific("Diferencia",this.getHoja(true));
            evaluator.evaluateFormulaCell(c5);
            Cell c6 = getCellSpecific("TotalEfectivo",this.getHoja(true));
            evaluator.evaluateFormulaCell(c6);
            Cell c7 = getCellSpecific("TotalDeposito",this.getHoja(true));
            evaluator.evaluateFormulaCell(c7);
        }
        public int getHoja(boolean ofToday){
            String fechaHoy = this.todayDate();
            String fecha = this.tomorrowDate();
            int indexHoja = 0;
            //Hoja de mañana
            Date date = new Date();
            if(!ofToday) {
                //Cambia de mes?
                System.out.println("Comparison: "+fechaHoy.split("-")[1] + " - "+fecha.split("-")[1]);
                if(fechaHoy.split("-")[1].equals(fecha.split("-")[1])){ //No cambia de mes
                    indexHoja = Integer.parseInt(fecha.split("-")[0]) - 1;
                }else if(date.getDay()==6)//es sabado, ocupamos el lunes
                    indexHoja = 51;
                else{ //se acabo el mes.
                    indexHoja=50;
                }
            }else{
                indexHoja = Integer.parseInt(fechaHoy.split("-")[0]) - 1;
            }
            return indexHoja;
        }
        
        

    /*
    Obtiene la fecha de hoy.
    */
    public String todayDate(){
        return this.FechaHoy;
    }
     /*
    Obtiene la fecha de mañana de la computadora.
    */
     public String tomorrowDate(){
        DateFormat format = new SimpleDateFormat("dd-MM-yyyy");
        Date someDate = new Date();
	Date newDate = new Date(someDate.getTime() + TimeUnit.DAYS.toMillis( 1 ));
        return format.format(newDate); 
    }
     public String afterSundayDate(){
        DateFormat format = new SimpleDateFormat("dd-MM-yyyy");
        Date someDate = new Date();
	Date newDate = new Date(someDate.getTime() + TimeUnit.DAYS.toMillis( 2 ));
        return format.format(newDate); 
    }
}
        
 
	