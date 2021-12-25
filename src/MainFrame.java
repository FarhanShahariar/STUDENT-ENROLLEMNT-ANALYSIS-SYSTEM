

import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainFrame extends javax.swing.JFrame 
{
    JFileChooser fileChooser = null;
    static boolean c = false;
    static FileInputStream fis;
    static int x = 0, y = 0;
    
    public MainFrame() 
    {
        initComponents();
        this.setLocationRelativeTo(null);
        
        //fetchSemester();
        //fetchYear();
    }
    
    public void generateEnrollmentReport()
    {
        String FILE = "Enrollment Report.pdf";
        
        try 
        {
            Document document = new Document();
            PdfWriter.getInstance(document, new FileOutputStream(FILE));
            document.open();
            // Left
            Paragraph paragraph = new Paragraph("This is right aligned text");
            paragraph.setAlignment(Element.ALIGN_RIGHT);
            document.add(paragraph);
            // Centered
            paragraph = new Paragraph("This is centered text");
            paragraph.setAlignment(Element.ALIGN_CENTER);
            document.add(paragraph);
            // Left
            paragraph = new Paragraph("This is left aligned text");
            paragraph.setAlignment(Element.ALIGN_LEFT);
            document.add(paragraph);
            // Left with indentation
            paragraph = new Paragraph(
                    "This is left aligned text with indentation");
            paragraph.setAlignment(Element.ALIGN_LEFT);
            paragraph.setIndentationLeft(50);
            document.add(paragraph);

            document.close();
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
    }
    
    /*public void generateRevenueReport(String par, String par1)
    {
        String FILE = "Revenue Report.pdf";
        
        try 
        {
            Document document = new Document();
            PdfWriter.getInstance(document, new FileOutputStream(FILE));
            document.setPageSize(A4);
            document.setMargins(2, 2, 2, 2);
            
            document.open();
            
            Image img = Image.getInstance("logo.jpg");
            img.setAlignment(Element.ALIGN_CENTER);
            
            document.add(img);

            document.close();
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
    }*/
    
    public boolean getExistingData(String semester, String year)
    {
        boolean exist = false;
        try 
        {
            String connectionUrl = "jdbc:sqlserver://localhost:1433;databaseName=SEAS;user=sa;password=123321";

            Connection con = DriverManager.getConnection(connectionUrl);
            Statement st = con.createStatement();
            ResultSet rs = st.executeQuery("SELECT distinct semester FROM excelTable where year = '"+year+"'");
            
            while (rs.next())
            {
                String sem = rs.getString(1);
                if(sem.equals(semester))
                {
                    exist = true;
                }
            }
        }
        catch (SQLException ex) 
        {
            JOptionPane.showMessageDialog(null, ex.getMessage());
        }
        
        return exist;
    }

    public String getFacultyId(String idd)
    {
        String id = "";
        String str = idd;
        id = str.substring(0, 4);
        return id;
    }
    
    public String getFacultyName(String idd)
    {
        String id = "";
        String str = idd;
        id = str.substring(5, idd.length()-1);
        return id;
    }
    
    /*public String getFacultyName(String name)
    {
        String nam = "";
        String str = name;
        name = str.substring(5, 3);
        return id;
    }*/
    
    public void insertIntoOfferCourse() throws SQLException
    {
        String connectionUrl = "jdbc:sqlserver://localhost:1433;databaseName=SEAS;user=sa;password=123321";
        Connection con = DriverManager.getConnection(connectionUrl);
        
        Statement st = con.createStatement();
        ResultSet rs = st.executeQuery("SELECT course_id, section, capacity, enrolled, Room_id, FACULTY_FULL_NAME, semester, year, SCHOOL_TITLE FROM excelTable "
                + "WHERE SEMESTER = '" + jComboBox1.getSelectedItem().toString() + "' AND YEAR = '" + yearTxt.getText() + "'");
        
        con = DriverManager.getConnection(connectionUrl);
                    PreparedStatement pst = con.prepareStatement("DELETE FROM OFFERED_COURSE where year = '" + yearTxt.getText() + "' and SEMESTER = '" + jComboBox1.getSelectedItem().toString() + "'");
                    pst.executeUpdate();
        
        while (rs.next())
        {
            String sql = "INSERT INTO OFFERED_COURSE (COURSE_ID,SECTION,CAPACITY,ENROLLED,ROOM_ID,FACULTY_ID,SEMESTER,YEAR,SCHOOL_ID) "
                    + "VALUES('" + rs.getString(1) + "','" + rs.getString(2) + "','" + rs.getString(3) + "','" + rs.getString(4) + "', '" + rs.getString(5) + "', '" + getFacultyId(rs.getString(6)) + "', "
                    + "'" + rs.getString(7) + "','" + rs.getString(8) + "', '" + rs.getString(9) + "')";
            
            System.out.println(sql);
            
            con = DriverManager.getConnection(connectionUrl);
             pst = con.prepareStatement(sql);
            pst.executeUpdate();
        }
        JOptionPane.showMessageDialog(rootPane, "System Refreshed");
    }
    
    public static void insertIntoROOM() throws SQLException
    {
        String connectionUrl = "jdbc:sqlserver://localhost:1433;databaseName=SEAS;user=sa;password=123321";
        Connection con = DriverManager.getConnection(connectionUrl);
        
        Statement st = con.createStatement();
        ResultSet rs = st.executeQuery("SELECT Distinct Room_id, ROOM_CAPACITY FROM excelTable "
                + "WHERE SEMESTER = 'Spring' AND YEAR = '2021'");
        
        con = DriverManager.getConnection(connectionUrl);
                    PreparedStatement pst = con.prepareStatement("DELETE FROM ROOM");
                    pst.executeUpdate();
        
        while (rs.next())
        {
            String sql = "INSERT INTO ROOM (ROOM_ID, ROOMSIZE) VALUES('" + rs.getString(1) + "','" + rs.getString(2) + "')";
                    System.out.println(sql);

                    con = DriverManager.getConnection(connectionUrl);
                    pst = con.prepareStatement(sql);
                    pst.executeUpdate();
                    
                    System.out.println(sql);
        }
        JOptionPane.showMessageDialog(null, "System Refreshed");
    }
    
    public  void insertIntoFACULTY() throws SQLException
    {
        String connectionUrl = "jdbc:sqlserver://localhost:1433;databaseName=SEAS;user=sa;password=123321";
        Connection con = DriverManager.getConnection(connectionUrl);
        
        Statement st = con.createStatement();
        ResultSet rs = st.executeQuery("SELECT Distinct FACULTY_FULL_NAME FROM excelTable");
        
        con = DriverManager.getConnection(connectionUrl);
                    PreparedStatement pst = con.prepareStatement("DELETE FROM FACULTY");
                    pst.executeUpdate();
        
        while (rs.next())
        {
            String sql = "INSERT INTO FACULTY (FACULTY_ID, FACULTY_FUL_NAME) VALUES('" + getFacultyId(rs.getString(1)) + "','" + getFacultyName(rs.getString(1)) + "')";
                    System.out.println(sql);

                    con = DriverManager.getConnection(connectionUrl);
                    pst = con.prepareStatement(sql);
                    pst.executeUpdate();
                    
                    System.out.println(sql);
        }
        JOptionPane.showMessageDialog(null, "System Refreshed");
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        jButton1 = new javax.swing.JButton();
        xlFileLocationtxt = new javax.swing.JTextField();
        jPanel2 = new javax.swing.JPanel();
        yearTxt = new javax.swing.JTextField();
        jComboBox1 = new javax.swing.JComboBox<>();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jButton2 = new javax.swing.JButton();
        jPanel3 = new javax.swing.JPanel();
        jComboBox3 = new javax.swing.JComboBox<>();
        jLabel7 = new javax.swing.JLabel();
        jButton4 = new javax.swing.JButton();
        jLabel8 = new javax.swing.JLabel();
        jLabel9 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        jComboBox5 = new javax.swing.JComboBox<>();
        jComboBox10 = new javax.swing.JComboBox<>();
        jComboBox11 = new javax.swing.JComboBox<>();
        jPanel5 = new javax.swing.JPanel();
        jComboBox4 = new javax.swing.JComboBox<>();
        jLabel10 = new javax.swing.JLabel();
        jButton5 = new javax.swing.JButton();
        jLabel13 = new javax.swing.JLabel();
        jLabel16 = new javax.swing.JLabel();
        jLabel17 = new javax.swing.JLabel();
        jComboBox12 = new javax.swing.JComboBox<>();
        jComboBox13 = new javax.swing.JComboBox<>();
        jComboBox14 = new javax.swing.JComboBox<>();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setMaximumSize(new java.awt.Dimension(672, 473));
        setResizable(false);

        jPanel1.setBackground(new java.awt.Color(51, 153, 255));

        jLabel1.setBackground(new java.awt.Color(0, 204, 204));
        jLabel1.setFont(new java.awt.Font("Times New Roman", 1, 18)); // NOI18N
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("STUDENT ENROLLMENT ANALYSIS SYSTEM");
        jLabel1.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 255, 255)));
        jLabel1.setOpaque(true);

        jButton1.setText("Read File");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        xlFileLocationtxt.setEditable(false);
        xlFileLocationtxt.setBackground(new java.awt.Color(255, 255, 102));
        xlFileLocationtxt.setFont(new java.awt.Font("Times New Roman", 0, 14)); // NOI18N
        xlFileLocationtxt.setForeground(new java.awt.Color(102, 102, 102));
        xlFileLocationtxt.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                xlFileLocationtxtMouseClicked(evt);
            }
        });
        xlFileLocationtxt.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                xlFileLocationtxtActionPerformed(evt);
            }
        });

        jPanel2.setBorder(javax.swing.BorderFactory.createTitledBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)), "Save Data"));
        jPanel2.setOpaque(false);

        yearTxt.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        yearTxt.setBorder(javax.swing.BorderFactory.createCompoundBorder(null, javax.swing.BorderFactory.createEmptyBorder(1, 6, 1, 1)));

        jComboBox1.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Spring", "Summer", "Autumn" }));
        jComboBox1.setToolTipText("");
        jComboBox1.setBorder(null);

        jLabel2.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel2.setForeground(new java.awt.Color(255, 255, 255));
        jLabel2.setText("Semester");

        jLabel3.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel3.setForeground(new java.awt.Color(255, 255, 255));
        jLabel3.setText("Year");

        jButton2.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jButton2.setText("Save");
        jButton2.setBorder(null);
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel2)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 113, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(25, 25, 25)
                .addComponent(jLabel3)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(yearTxt, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 89, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel3)
                    .addComponent(yearTxt, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap())
        );

        jPanel3.setBorder(javax.swing.BorderFactory.createTitledBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)), "Enrollment Report"));
        jPanel3.setOpaque(false);

        jComboBox3.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jComboBox3.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "2021", "2020" }));
        jComboBox3.setToolTipText("");
        jComboBox3.setOpaque(false);

        jLabel7.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel7.setForeground(new java.awt.Color(255, 255, 255));
        jLabel7.setText("Year");

        jButton4.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jButton4.setText("Generate");
        jButton4.setOpaque(false);
        jButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton4ActionPerformed(evt);
            }
        });

        jLabel8.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel8.setForeground(new java.awt.Color(255, 255, 255));
        jLabel8.setText("Year");

        jLabel9.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel9.setForeground(new java.awt.Color(255, 255, 255));
        jLabel9.setText("Semester");

        jLabel6.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel6.setForeground(new java.awt.Color(255, 255, 255));
        jLabel6.setText("Semester");

        jComboBox5.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jComboBox5.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "2021", "2020" }));
        jComboBox5.setToolTipText("");
        jComboBox5.setOpaque(false);

        jComboBox10.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jComboBox10.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Spring", "Summer", "Autumn" }));
        jComboBox10.setToolTipText("");
        jComboBox10.setBorder(null);

        jComboBox11.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jComboBox11.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Spring", "Summer", "Autumn" }));
        jComboBox11.setToolTipText("");
        jComboBox11.setBorder(null);

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addComponent(jLabel6)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jComboBox10, javax.swing.GroupLayout.PREFERRED_SIZE, 113, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(25, 25, 25)
                        .addComponent(jLabel7))
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addComponent(jLabel9)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jComboBox11, javax.swing.GroupLayout.PREFERRED_SIZE, 113, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jLabel8)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jComboBox3, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jComboBox5, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(20, 20, 20)
                .addComponent(jButton4)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                    .addComponent(jButton4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel7)
                            .addComponent(jComboBox3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel6)
                            .addComponent(jComboBox10, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(26, 26, 26)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel9)
                            .addComponent(jLabel8)
                            .addComponent(jComboBox5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jComboBox11, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addGap(20, 20, 20))
        );

        jPanel5.setBorder(javax.swing.BorderFactory.createTitledBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)), "Revenue Report"));
        jPanel5.setOpaque(false);

        jComboBox4.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jComboBox4.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "2021", "2020" }));
        jComboBox4.setToolTipText("");
        jComboBox4.setOpaque(false);

        jLabel10.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel10.setForeground(new java.awt.Color(255, 255, 255));
        jLabel10.setText("Year");

        jButton5.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jButton5.setText("Generate");
        jButton5.setOpaque(false);
        jButton5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton5ActionPerformed(evt);
            }
        });

        jLabel13.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel13.setForeground(new java.awt.Color(255, 255, 255));
        jLabel13.setText("Year");

        jLabel16.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel16.setForeground(new java.awt.Color(255, 255, 255));
        jLabel16.setText("Semester");

        jLabel17.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel17.setForeground(new java.awt.Color(255, 255, 255));
        jLabel17.setText("Semester");

        jComboBox12.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jComboBox12.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "2021", "2020" }));
        jComboBox12.setToolTipText("");
        jComboBox12.setOpaque(false);

        jComboBox13.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jComboBox13.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Spring", "Summer", "Autumn" }));
        jComboBox13.setToolTipText("");
        jComboBox13.setBorder(null);

        jComboBox14.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jComboBox14.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Spring", "Summer", "Autumn" }));
        jComboBox14.setToolTipText("");
        jComboBox14.setBorder(null);

        javax.swing.GroupLayout jPanel5Layout = new javax.swing.GroupLayout(jPanel5);
        jPanel5.setLayout(jPanel5Layout);
        jPanel5Layout.setHorizontalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addComponent(jLabel17)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jComboBox13, javax.swing.GroupLayout.PREFERRED_SIZE, 113, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(25, 25, 25)
                        .addComponent(jLabel10))
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addComponent(jLabel16)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jComboBox14, javax.swing.GroupLayout.PREFERRED_SIZE, 113, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jLabel13)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jComboBox4, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jComboBox12, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(20, 20, 20)
                .addComponent(jButton5)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel5Layout.setVerticalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                    .addComponent(jButton5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel10)
                            .addComponent(jComboBox4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel17)
                            .addComponent(jComboBox13, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(26, 26, 26)
                        .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel16)
                            .addComponent(jLabel13)
                            .addComponent(jComboBox12, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jComboBox14, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addGap(20, 20, 20))
        );

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(10, 10, 10)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                    .addComponent(jPanel3, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanel2, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(xlFileLocationtxt, javax.swing.GroupLayout.PREFERRED_SIZE, 528, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 13, Short.MAX_VALUE)
                        .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 111, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(xlFileLocationtxt, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(20, 20, 20)
                .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(20, 20, 20)
                .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(20, 20, 20)
                .addComponent(jPanel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(20, 20, 20))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        try 
        {
            XSSFWorkbook workbook = new XSSFWorkbook(xlFileLocationtxt.getText());
            XSSFSheet s = workbook.getSheetAt(0);
            XSSFRow r = s.getRow(0);
            int col = r.getLastCellNum();
            if(workbook.getNumberOfSheets() != 1 || col != 15)
            {
                JOptionPane.showMessageDialog(rootPane, "Excel file sheets is not formated properly!");
                c = false;
                System.out.println("Format");
            }
            else
            {
                JOptionPane.showMessageDialog(rootPane, "Excel File Loaded");
                c = true;
            }
        } 
        catch(IOException ex) 
        {
            JOptionPane.showMessageDialog(rootPane, ex.getMessage());
        }
    }//GEN-LAST:event_jButton1ActionPerformed

    private void xlFileLocationtxtActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_xlFileLocationtxtActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_xlFileLocationtxtActionPerformed

    private void xlFileLocationtxtMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_xlFileLocationtxtMouseClicked
        fileChooser = new JFileChooser();
        fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("MS Office Documents","xlsx"));
        fileChooser.setAcceptAllFileFilterUsed(true);
        int result = fileChooser.showOpenDialog(null);        
        if(result == JFileChooser.APPROVE_OPTION) 
        {
            String ex ;
            File selectedFile = fileChooser.getSelectedFile();
            String name = selectedFile.getName();
            ex = name.substring(name.lastIndexOf("."));
            if(ex.equals(".xlsx"))
            {
                xlFileLocationtxt.setText(selectedFile.getAbsolutePath());
                selectedFile = new File(xlFileLocationtxt.getText());
                try 
                {
                    fis = new FileInputStream(selectedFile.getAbsolutePath());
                } 
                catch (FileNotFoundException ex1) 
                {
                    JOptionPane.showMessageDialog(rootPane, ex1.getMessage());
                }
            }
        }
    }//GEN-LAST:event_xlFileLocationtxtMouseClicked

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        if(yearTxt.getText().isEmpty() || c == false)
        {
            JOptionPane.showMessageDialog(rootPane, "Enter Year");
        }
        else if(getExistingData(jComboBox1.getSelectedItem().toString(), yearTxt.getText()))
        {
            JOptionPane.showMessageDialog(rootPane, "Data Exist");
        }
        else
        {
            //if(getExistingData(semester,year))
            try 
            { 
                
                XSSFWorkbook wb = new XSSFWorkbook(xlFileLocationtxt.getText());
                XSSFSheet sheet = wb.getSheetAt(0);
                int row = sheet.getLastRowNum();
                int col = 11;//r.getLastCellNum();

                System.out.println(row);
                System.out.println(col);

                for (int i = 1; i <= row; i++) {
                    DataFormatter df = new DataFormatter();
                    StringBuilder sb = new StringBuilder();
                    XSSFRow getRow = sheet.getRow(i);
                    CellType type;

                    for (int j = 0; j <= col; j++) {
                        Cell cell = getRow.getCell(j);

                        if (cell == null || cell.getCellType() == CellType.BLANK) {
                            type = CellType.BLANK;
                        } else {
                            type = cell.getCellType();
                        }

                        switch (type) {
                            case NUMERIC:
                                sb.append(df.formatCellValue(cell));

                                if (j != col) 
                                {
                                    sb.append(",");
                                }
                                break;
                            case STRING:
                                String k = cell.getStringCellValue();
                                if (k.contains("'")) {
                                    k = k.replace("'", "''");
                                }
                                sb.append("'").append(k).append("'");
                                if (j != col) {
                                    sb.append(",");
                                }
                                break;
                            case BLANK:
                                sb.append("NULL");
                                if (j != col) {
                                    sb.append(",");
                                }
                                break;
                            default:
                                break;
                        }
                    }
                    String sql = "INSERT INTO excelTable (SEMESTER,YEAR,SCHOOL_TITLE,COURSE_ID,COFFERED_COURSE,SECTION,CREDIT_HOUR,CAPACITY,ENROLLED,ROOM_ID,ROOM_CAPACITY,BLOCKED,COURSE_TITLE,"
                            + "FACULTY_FULL_NAME) VALUES('"+jComboBox1.getSelectedItem().toString()+"',"+yearTxt.getText() + ","+sb.toString() + ")";
                    
                    System.out.println(sql);
                    
                    String connectionUrl = "jdbc:sqlserver://localhost:1433;databaseName=SEAS;user=sa;password=123321";

                    Connection con = DriverManager.getConnection(connectionUrl);
                    PreparedStatement pst = con.prepareStatement(sql);
                    pst.executeUpdate();
                    
                    
                }
                
                JOptionPane.showMessageDialog(rootPane, "Data Uploaded");
                insertIntoROOM();
                insertIntoFACULTY();
                insertIntoOfferCourse();
                    
            }
            catch (IOException e) 
            {
                JOptionPane.showMessageDialog(rootPane, e.getMessage());
            } catch (SQLException ex) 
            {
                Logger.getLogger(MainFrame.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }//GEN-LAST:event_jButton2ActionPerformed

    private void jButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton4ActionPerformed
        NewClass.generateRevenueReport();
    }//GEN-LAST:event_jButton4ActionPerformed

    private void jButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton5ActionPerformed
        //generateRevenueReport();        // TODO add your handling code here:
    }//GEN-LAST:event_jButton5ActionPerformed

   
    public static void main(String args[]) 
    {
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try 
        {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) 
            {
                if ("Nimbus".equals(info.getName())) 
                {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new MainFrame().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton4;
    private javax.swing.JButton jButton5;
    public static javax.swing.JComboBox<String> jComboBox1;
    public static javax.swing.JComboBox<String> jComboBox10;
    public static javax.swing.JComboBox<String> jComboBox11;
    private javax.swing.JComboBox<String> jComboBox12;
    private javax.swing.JComboBox<String> jComboBox13;
    private javax.swing.JComboBox<String> jComboBox14;
    public static javax.swing.JComboBox<String> jComboBox3;
    private javax.swing.JComboBox<String> jComboBox4;
    public static javax.swing.JComboBox<String> jComboBox5;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel16;
    private javax.swing.JLabel jLabel17;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel5;
    private javax.swing.JTextField xlFileLocationtxt;
    public static javax.swing.JTextField yearTxt;
    // End of variables declaration//GEN-END:variables
}
