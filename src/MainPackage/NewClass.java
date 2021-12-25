package MainPackage;


import static MainPackage.MainFrame.jComboBox1;
import static MainPackage.MainFrame.yearTxt;
import com.itextpdf.awt.PdfGraphics2D;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Chunk;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Font.FontFamily;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;


public class NewClass 
{
     public static DefaultCategoryDataset createDataset() throws SQLException {

        String semester;
        String SBE;
        String SETS;
        String SELS;
        String SLASS;
        String SPPH;
        String Total;
        String Change;

        DefaultCategoryDataset dataset = new DefaultCategoryDataset();

        String connectionUrl = "jdbc:sqlserver://localhost:1433;databaseName=SEAS;user=sa;password=123321";
            Connection con = DriverManager.getConnection(connectionUrl);
        
            Statement st = con.createStatement();
            ResultSet rs = st.executeQuery("Select SEMESTER, year, SCHOOL_id, sum(Credit*ENROLLED) FROM OFFERED_COURSE where  year >= 2021 and  year <= 2021 group by SEMESTER, YEAR, "
                    + "SCHOOL_id order by YEAR");
            while(rs.next()){
            
        dataset.addValue(Integer.valueOf(rs.getString(4)), rs.getString(3) , rs.getString(1)+" "+rs.getString(2));
            }
        
        /*dataset.addValue(150, series1, "2016-12-20");
        dataset.addValue(100, series1, "2016-12-21");
        dataset.addValue(210, series1, "2016-12-22");
        dataset.addValue(240, series1, "2016-12-23");
        dataset.addValue(195, series1, "2016-12-24");
        dataset.addValue(245, series1, "2016-12-25");

        dataset.addValue(150, series2, "2016-12-19");
        dataset.addValue(130, series2, "2016-12-20");
        dataset.addValue(95, series2, "2016-12-21");
        dataset.addValue(195, series2, "2016-12-22");
        dataset.addValue(200, series2, "2016-12-23");
        dataset.addValue(180, series2, "2016-12-24");
        dataset.addValue(230, series2, "2016-12-25");*/

        return dataset;
    }
     
    public static JFreeChart lineChart() throws SQLException 
    {
	DefaultCategoryDataset dataset = createDataset();  
        
        JFreeChart chart = ChartFactory.createLineChart( "Revenue Trend of the schools", // Chart title  
        "Semester", // X-Axis Label  
        "Credit Hour", // Y-Axis Label  
        dataset ,PlotOrientation.VERTICAL,
         true,true,false
        );  

      
        return chart;
    }
    
    
    public static void generateRevenueReport()
    {
        String FILE = "Revenue Report.pdf";
        
        try 
        {
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(FILE));
            document.setMargins(45f, 45f, 50f, 50f);
            
            Font myFont = new Font(FontFamily.TIMES_ROMAN); 
            
            //myContentStyle.setStyle("bold");
            //myContentStyle.setColor(255,0,0);
            //mySizeSpecification.setSize(20);
            
            document.open();
            
            Image img = Image.getInstance("logo.jpg");
            img.setAlignment(Element.ALIGN_CENTER);
            
            document.add(img);
            
            myFont.setSize(27);
            myFont.setStyle("bold");
            Paragraph paragraph = new Paragraph("Independent University, Bangladesh", myFont);
            paragraph.setAlignment(Element.ALIGN_CENTER);
            document.add(paragraph);
            
            myFont.setSize(19);
            myFont.setStyle("normal");
            paragraph = new Paragraph("Revenue of IUB and SETS", myFont);
            paragraph.setAlignment(Element.ALIGN_CENTER);
            document.add(paragraph);
                    
            myFont.setSize(15);
            paragraph = new Paragraph("Period: ", myFont);
            paragraph.setAlignment(Element.ALIGN_CENTER);
            document.add(paragraph);
            
            myFont.setSize(15);
            paragraph = new Paragraph("Source: IRAS (online data)", myFont);
            paragraph.setAlignment(Element.ALIGN_CENTER);
            document.add(paragraph);
            
            document.newPage();
            myFont = new Font(FontFamily.TIMES_ROMAN); 
            
            myFont.setSize(12);
            myFont.setStyle("bold");
            paragraph = new Paragraph("Revenue:", myFont);
            document.add(paragraph);
            
            myFont = new Font(FontFamily.TIMES_ROMAN); 
            myFont.setSize(12);
            paragraph = new Paragraph("This report assumes that revenue is sum of total credit hours the students have enrolled in for a particular semester. "
                    + "The revenue has been computed using the tally sheet where the course details along with credit hours and no of enrollment is provided. \n" +
                        "Revenue of IUB", myFont);
            document.add(paragraph);
            
            document.add(new Chunk(" "));
            myFont.setStyle("bold");
            paragraph = new Paragraph("The raw data of", myFont);
            paragraph.setSpacingAfter(10f);
            document.add(paragraph);
            
            // Creating a table     
            myFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 9);//new Font(FontFamily.TIMES_ROMAN);
            PdfPTable table = new PdfPTable(43);
            table.setHorizontalAlignment(Element.ALIGN_LEFT);
            table.setWidthPercentage(100);
            
            PdfPCell cell1 = new PdfPCell(new Paragraph("Semester",myFont));
            PdfPCell cell2 = new PdfPCell(new Paragraph("SBE",myFont));
            PdfPCell cell3 = new PdfPCell(new Paragraph("SETS",myFont));
            PdfPCell cell4 = new PdfPCell(new Paragraph("SELS",myFont));
            PdfPCell cell5 = new PdfPCell(new Paragraph("SLASS",myFont));
            PdfPCell cell6 = new PdfPCell(new Paragraph("SPPH",myFont));
            PdfPCell cell7 = new PdfPCell(new Paragraph("Total",myFont));
            PdfPCell cell8 = new PdfPCell(new Paragraph("Change %",myFont));

            cell1.setHorizontalAlignment(Element.ALIGN_LEFT);
            cell1.setBackgroundColor(new BaseColor(237, 125, 49));
            cell1.setColspan(8);
            table.addCell(cell1);
            
            cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell2.setBackgroundColor(new BaseColor(237, 125, 49));
            cell2.setColspan(5);
            table.addCell(cell2);
            
            cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell3.setBackgroundColor(new BaseColor(237, 125, 49));
            cell3.setColspan(5);
            table.addCell(cell3);
            
            cell4.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell4.setBackgroundColor(new BaseColor(237, 125, 49));
            cell4.setColspan(5);
            table.addCell(cell4);
            
            cell5.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell5.setBackgroundColor(new BaseColor(237, 125, 49));
            cell5.setColspan(5);
            table.addCell(cell5);
            
            cell6.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell6.setBackgroundColor(new BaseColor(237, 125, 49));
            cell6.setColspan(5);
            table.addCell(cell6);
            
            cell7.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell7.setBackgroundColor(new BaseColor(237, 125, 49));
            cell7.setColspan(5);
            table.addCell(cell7);
            
            cell8.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell8.setBackgroundColor(new BaseColor(237, 125, 49));
            cell8.setColspan(5);
            table.addCell(cell8);
            
            
            List<productBean> list = new ArrayList<>();
            
            String connectionUrl = "jdbc:sqlserver://localhost:1433;databaseName=SEAS;user=sa;password=123321";
            Connection con = DriverManager.getConnection(connectionUrl);
        
            Statement st = con.createStatement();
            ResultSet rs = st.executeQuery("Select SEMESTER, year, SCHOOL_id, sum(Credit*ENROLLED) FROM OFFERED_COURSE where  year >= 2021 and  year <= 2021 group by SEMESTER, YEAR, "
                    + "SCHOOL_id order by YEAR");
            
            list.add(new productBean("Semester", "SBE", "SETS", "SELS", "SLASS", "SPPH", "Total", "Change"));
            
            System.out.println(list.toString());
            
            Iterator itr = list.iterator();  
                //traversing elements of ArrayList object  
                while(itr.hasNext()){  
                  productBean p = (productBean)itr.next();  
                  System.out.println(" "+p.Change+" "+p.SBE);  
                }  

            //for (int i = 0; i < 100; i++) 
            
            int i = 0;
            while(rs.next())
            {
                
            }
            
            
            
            
            {
                String prod = "", bran = "", sqnt = "", rat = "", tota = "", suplr = "", detl = "", mod = "", war = "";

                myFont = FontFactory.getFont(FontFactory.HELVETICA, 8);

                PdfPCell semester1 = new PdfPCell(new Phrase(prod, myFont));
                semester1.setBackgroundColor(new BaseColor(237, 125, 49));
                semester1.setColspan(8);
                semester1.setHorizontalAlignment(Element.ALIGN_LEFT);
                semester1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                table.addCell(semester1);
                
                int r , g, b;
                if(i % 2 == 0)
                {
                    r = 194; g = 213; b = 215;//, ,  
                }
                else
                {
                    r = 178; g = 195; b = 197;
                }
                
                PdfPCell Qnty1 = new PdfPCell(new Phrase(bran, myFont));
                Qnty1.setColspan(5);
                Qnty1.setHorizontalAlignment(Element.ALIGN_CENTER);
                Qnty1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                Qnty1.setBackgroundColor(new BaseColor(r, g, b));
                table.addCell(Qnty1);
                
                PdfPCell model = new PdfPCell(new Phrase(mod, myFont));
                model.setColspan(5);
                model.setHorizontalAlignment(Element.ALIGN_CENTER);
                model.setVerticalAlignment(Element.ALIGN_MIDDLE);
                model.setBackgroundColor(new BaseColor(r, g, b));
                table.addCell(model);
                
                PdfPCell warn = new PdfPCell(new Phrase(war, myFont));
                warn.setColspan(5);
                warn.setHorizontalAlignment(Element.ALIGN_CENTER);
                warn.setVerticalAlignment(Element.ALIGN_MIDDLE);
                warn.setBackgroundColor(new BaseColor(r, g, b));
                table.addCell(warn);
                
                PdfPCell Rate1 = new PdfPCell(new Phrase(sqnt, myFont));
                Rate1.setColspan(5);
                Rate1.setHorizontalAlignment(Element.ALIGN_CENTER);
                Rate1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                Rate1.setBackgroundColor(new BaseColor(r, g, b));
                table.addCell(Rate1);
                
                PdfPCell Discount1 = new PdfPCell(new Phrase(rat, myFont));
                Discount1.setColspan(5);
                Discount1.setHorizontalAlignment(Element.ALIGN_CENTER);
                Discount1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                Discount1.setBackgroundColor(new BaseColor(r, g, b));
                table.addCell(Discount1);
                
                PdfPCell Total1 = new PdfPCell(new Phrase(tota, myFont));
                Total1.setColspan(5);
                Total1.setHorizontalAlignment(Element.ALIGN_CENTER);
                Total1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                Total1.setBackgroundColor(new BaseColor(r, g, b));
                table.addCell(Total1);
                
                PdfPCell su = new PdfPCell(new Phrase(suplr, myFont));
                su.setColspan(5);
                su.setHorizontalAlignment(Element.ALIGN_CENTER);
                su.setVerticalAlignment(Element.ALIGN_MIDDLE);
                su.setBackgroundColor(new BaseColor(r, g, b));
                table.addCell(su);
                
                i++;
            }

            document.add(table); 
            
            
            PdfContentByte cb = writer.getDirectContent();
            float width = PageSize.A4.getWidth();
            float height = PageSize.A4.getHeight() / 2;
            // Pie chart
            PdfTemplate pie = cb.createTemplate(width, height);
            Graphics2D g2d1 = new PdfGraphics2D(pie, width, height);
            Rectangle2D r2d1 = new Rectangle2D.Double(0, 0, width, height);
            lineChart().draw(g2d1, r2d1);
            g2d1.dispose();
            cb.addTemplate(pie, 0, height);
            
            
           /* 
            document.newPage();
            myFont = FontFactory.getFont(FontFactory.TIMES_BOLD, 12);
            paragraph = new Paragraph("Revenue in Engineering School:", myFont);
            paragraph.setSpacingAfter(10f);
            document.add(paragraph);
            
            // Creating a table     
            myFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 9);//new Font(FontFamily.TIMES_ROMAN);
            PdfPTable table2 = new PdfPTable(48);
            table2.setHorizontalAlignment(Element.ALIGN_LEFT);
            table2.setWidthPercentage(100);
            
            PdfPCell cell11 = new PdfPCell(new Paragraph("Semester",myFont));
            PdfPCell cell22 = new PdfPCell(new Paragraph("SBE",myFont));
            PdfPCell cell33 = new PdfPCell(new Paragraph("SETS",myFont));
            PdfPCell cell44 = new PdfPCell(new Paragraph("SELS",myFont));
            PdfPCell cell55 = new PdfPCell(new Paragraph("SLASS",myFont));
            PdfPCell cell66 = new PdfPCell(new Paragraph("SPPH",myFont));
            PdfPCell cell77 = new PdfPCell(new Paragraph("Total",myFont));
            PdfPCell cell88 = new PdfPCell(new Paragraph("Change %",myFont));
            PdfPCell cell99 = new PdfPCell(new Paragraph("Change %",myFont));

            cell11.setHorizontalAlignment(Element.ALIGN_LEFT);
            cell11.setBackgroundColor(new BaseColor(237, 125, 49));
            cell11.setColspan(8);
            table2.addCell(cell11);
            
            cell22.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell22.setBackgroundColor(new BaseColor(237, 125, 49));
            cell22.setColspan(5);
            table2.addCell(cell22);
            
            cell33.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell33.setBackgroundColor(new BaseColor(237, 125, 49));
            cell33.setColspan(5);
            table2.addCell(cell33);
            
            cell44.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell44.setBackgroundColor(new BaseColor(237, 125, 49));
            cell44.setColspan(5);
            table2.addCell(cell44);
            
            cell55.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell55.setBackgroundColor(new BaseColor(237, 125, 49));
            cell55.setColspan(5);
            table2.addCell(cell55);
            
            cell66.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell66.setBackgroundColor(new BaseColor(237, 125, 49));
            cell66.setColspan(5);
            table2.addCell(cell66);
            
            cell77.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell77.setBackgroundColor(new BaseColor(237, 125, 49));
            cell77.setColspan(5);
            table2.addCell(cell77);
            
            cell88.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell88.setBackgroundColor(new BaseColor(237, 125, 49));
            cell88.setColspan(5);
            table2.addCell(cell88);
            
            cell99.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell99.setBackgroundColor(new BaseColor(237, 125, 49));
            cell99.setColspan(5);
            table2.addCell(cell99);

            for (int i = 0; i < 100; i++) 
            {
                String prod = "555", bran = "", sqnt = "", rat = "", tota = "", suplr = "", detl = "", mod = "", war = "";

                myFont = FontFactory.getFont(FontFactory.HELVETICA, 8);

                PdfPCell semester1 = new PdfPCell(new Phrase(prod, myFont));
                semester1.setBackgroundColor(new BaseColor(237, 125, 49));
                semester1.setColspan(8);
                semester1.setHorizontalAlignment(Element.ALIGN_LEFT);
                semester1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                table2.addCell(semester1);
                
                int r , g, b;
                if(i % 2 == 0)
                {
                    r = 194; g = 213; b = 215;//, ,  
                }
                else
                {
                    r = 178; g = 195; b = 197;
                }
                
                PdfPCell Qnty1 = new PdfPCell(new Phrase(bran, myFont));
                Qnty1.setColspan(5);
                Qnty1.setHorizontalAlignment(Element.ALIGN_CENTER);
                Qnty1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                Qnty1.setBackgroundColor(new BaseColor(r, g, b));
                table2.addCell(Qnty1);
                
                PdfPCell model = new PdfPCell(new Phrase(mod, myFont));
                model.setColspan(5);
                model.setHorizontalAlignment(Element.ALIGN_CENTER);
                model.setVerticalAlignment(Element.ALIGN_MIDDLE);
                model.setBackgroundColor(new BaseColor(r, g, b));
                table2.addCell(model);
                
                PdfPCell warn = new PdfPCell(new Phrase(war, myFont));
                warn.setColspan(5);
                warn.setHorizontalAlignment(Element.ALIGN_CENTER);
                warn.setVerticalAlignment(Element.ALIGN_MIDDLE);
                warn.setBackgroundColor(new BaseColor(r, g, b));
                table2.addCell(warn);
                
                PdfPCell Rate1 = new PdfPCell(new Phrase(sqnt, myFont));
                Rate1.setColspan(5);
                Rate1.setHorizontalAlignment(Element.ALIGN_CENTER);
                Rate1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                Rate1.setBackgroundColor(new BaseColor(r, g, b));
                table2.addCell(Rate1);
                
                PdfPCell Discount1 = new PdfPCell(new Phrase(rat, myFont));
                Discount1.setColspan(5);
                Discount1.setHorizontalAlignment(Element.ALIGN_CENTER);
                Discount1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                Discount1.setBackgroundColor(new BaseColor(r, g, b));
                table2.addCell(Discount1);
                
                PdfPCell Total1 = new PdfPCell(new Phrase(tota, myFont));
                Total1.setColspan(5);
                Total1.setHorizontalAlignment(Element.ALIGN_CENTER);
                Total1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                Total1.setBackgroundColor(new BaseColor(r, g, b));
                table2.addCell(Total1);
                
                PdfPCell su = new PdfPCell(new Phrase(suplr, myFont));
                su.setColspan(5);
                su.setHorizontalAlignment(Element.ALIGN_CENTER);
                su.setVerticalAlignment(Element.ALIGN_MIDDLE);
                su.setBackgroundColor(new BaseColor(r, g, b));
                table2.addCell(su);
                
                PdfPCell sup = new PdfPCell(new Phrase(detl, myFont));
                sup.setColspan(5);
                sup.setHorizontalAlignment(Element.ALIGN_CENTER);
                sup.setVerticalAlignment(Element.ALIGN_MIDDLE);
                sup.setBackgroundColor(new BaseColor(r, g, b));
                table2.addCell(sup);
            }

            document.add(table2); 
*/
            document.close();
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
    }
    
    public static String getFacultyId(String idd)
    {
        String id = "";
        String str = idd;
        id = str.substring(0, 4);
        return id;
    }
    
    public static String getFacultyName(String idd)
    {
        String id = "";
        String str = idd;
        id = str.substring(5, idd.length());
        return id;
    }
    
    public static void main(String args []) throws SQLException
    {
        generateRevenueReport();
  
        /*String id = "";
        String str = "4275-MS Tahsin F. Ara Nayna";
        id = str.substring(5, str.length()-1);
        System.out.println(id);
        */

        
    }
    
}


