/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Iterator;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;


public class NewServlet extends HttpServlet {

    /**
     * Processes requests for both HTTP <code>GET</code> and <code>POST</code>
     * methods.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
   

    
    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        doPost(request, response);
        
    }

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        PrintWriter pw=response.getWriter();
    String name=request.getParameter("name");
    String id=request.getParameter("id");
    String date=request.getParameter("date");
    String workDescp=request.getParameter("workDescription");
    String excelFilePath = "E:\\Book1.xlsx";  
        try {
            XSSFWorkbook workbook;
            try (FileInputStream inputStream = new FileInputStream(excelFilePath)) {
                workbook = new XSSFWorkbook(inputStream);
                XSSFSheet sheet = workbook.getSheet("Employee Data");
                XSSFFont font= workbook.createFont();
                CellStyle style=workbook.createCellStyle();
                int rowCount = sheet.getLastRowNum();
                if (rowCount==0){
                Row header = sheet.createRow(0);
                font.setBold(true);
                font.setFontName("Arial");
                font.setColor(IndexedColors.DARK_BLUE.getIndex());
                style.setFont(font);
                header.createCell(0).setCellValue("Employee name");
                sheet.autoSizeColumn(0);
                header.getCell(0).setCellStyle(style);
                header.createCell(1).setCellValue("Employee ID");
                sheet.autoSizeColumn(1);
                header.getCell(1).setCellStyle(style);
                header.createCell(2).setCellValue("Date");
                sheet.autoSizeColumn(2);
                header.getCell(2).setCellStyle(style);
                header.createCell(3).setCellValue("Work Description");
                sheet.autoSizeColumn(3);
                header.getCell(3).setCellStyle(style);
                
                }
                Row row = sheet.createRow(++rowCount);         
                ArrayList<String> arraylist=new ArrayList<String>();
                arraylist.add(name);
                arraylist.add(id);
                arraylist.add(date);
                arraylist.add(workDescp);
                int i=0;
                for(String s1:arraylist)
                {
                  Cell cell= row.createCell(i);
                  sheet.autoSizeColumn(i);
                  cell.setCellValue(s1);
                  i++;
                }
                
            }
            try (FileOutputStream outputStream = new FileOutputStream("E:\\Book1.xlsx")) {
                workbook.write(outputStream);
                workbook.close();
                System.out.println("Added Sccessfully"+name+"---"+id+"---"+date+"---"+workDescp);
                response.sendRedirect("index.html");
            }             
        } catch (IOException e) {
            e.printStackTrace();
        }
        
    }

    /**
     * Returns a short description of the servlet.
     *
     * @return a String containing servlet description
     */
    @Override
    public String getServletInfo() {
        return "Short description";
    }// </editor-fold>

}
