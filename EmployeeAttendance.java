package employeeattendance;

import com.google.zxing.EncodeHintType;
import com.google.zxing.NotFoundException;
import com.google.zxing.WriterException;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;
import java.awt.Color;
import java.awt.FlowLayout;
import java.awt.GridBagLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.*;
import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import javax.swing.JPasswordField;
import javax.swing.JTextField;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.regex.Matcher.*;
import java.util.regex.Pattern.*;
import java.util.regex.*;

public class EmployeeAttendance 
{
  public static boolean isValidEmailAddress(String email) {
           String ePattern = "^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\])|(([a-zA-Z\\-0-9]+\\.)+[a-zA-Z]{2,}))$";
           Pattern p = Pattern.compile(ePattern);
           Matcher m = p.matcher(email);
           return m.matches();
    }
  
  public static boolean isValidPhoneNumber(String phone){
  String ePattern="^[7-9][0-9]{9}$";
  Pattern p=Pattern.compile(ePattern);
  Matcher m=p.matcher(phone);
  return m.matches();
  }
  
    public static void main(String[] args) 
    {
       
        DateFormat df = new SimpleDateFormat("dd-MM-yyyy");
        String todayDate = df.format(new Date());
    
        File todayFile = new File(todayDate + ".xlsx");
        if(!todayFile.exists())
        {
            System.out.println("today file");
            File employeeDataFile = new File("EmployeeData.xlsx");
            
            if(employeeDataFile.exists())
            {
                try
                {
                    int rowCount,colCount;
                    int i,j;
                    FileInputStream data = new FileInputStream(employeeDataFile);
                    XSSFWorkbook wb = new XSSFWorkbook(data);
                    XSSFSheet st = wb.getSheetAt(0);
                    rowCount = st.getLastRowNum();
                   
                    XSSFWorkbook workbook = new XSSFWorkbook();
                    XSSFSheet sheet = workbook.createSheet();
                    Row row = sheet.createRow(0);
                    Cell cell = row.createCell(0);
                    cell.setCellValue((String) "EMPLOYEE ID" );
                    cell = row.createCell(1);
                    cell.setCellValue((String) "FIRST NAME" );
                    cell = row.createCell(2);
                    cell.setCellValue((String) "LAST NAME" );
                    cell = row.createCell(3);
                    cell.setCellValue((String) "EMAIL" );
                    cell = row.createCell(4);
                    cell.setCellValue((String) "PHONE" );
                    cell = row.createCell(5);
                    cell.setCellValue((String) "DESIGNATION" );
                    cell = row.createCell(6);
                    
                    cell.setCellValue((String) "TIME ENTRY" );
                    cell = row.createCell(7);
                    cell.setCellValue((String) "TIME EXIT" );
                    
                    for(i=1;i<=rowCount;i++)
                    {
                        Row r = st.getRow(i);    
                        row = sheet.createRow(i);
                        
                        String value;
                        for(j=0;j<6;j++)
                        {
                            cell = row.createCell(j);
                            value =  r.getCell(j).getStringCellValue();
                            cell.setCellValue((String)value);
                            
                        }
                        cell = row.createCell(j++);
                        cell.setCellValue((String) "Absent");
                        cell = row.createCell(j++);
                        cell.setCellValue((String) "Absent");
                    }
                    FileOutputStream out = new FileOutputStream(todayDate + ".xlsx");
                    workbook.write(out);
                }
                catch(Exception ex)
                {
                    System.out.println("Error occured");
                }
            }
        }
       
        final String userId="admin",password="admin";
       
        final JFrame f = new JFrame();
        try 
        {
            f.setContentPane(new JLabel(new ImageIcon(ImageIO.read(new File("loginImage.jpg")))));
        } 
        catch (IOException ex) 
        {
            Logger.getLogger(EmployeeAttendance.class.getName()).log(Level.SEVERE, null, ex);
        }
        final JTextField adminIdTextField;
        final JPasswordField passwordField;
        final JLabel adminIdLabel,passwordLabel,titleLabel;
        final JButton login,employee;
       
        titleLabel = new JLabel("Admin Login");
        titleLabel.setBounds(200,100,150,20);
        adminIdLabel = new JLabel("Admin Id");
        adminIdLabel.setBounds(100, 150, 100, 20);
        passwordLabel = new JLabel("Password");
        passwordLabel.setBounds( 100, 170, 100, 20);
        adminIdTextField = new JTextField();
        adminIdTextField.setBounds(200,150,100,20);
        passwordField = new JPasswordField();
        passwordField.setBounds(200,170,100,20);
        login = new JButton("Log in");
        login.setBounds(140,200,100,25);
       
        final JLabel errorMessage = new JLabel();
        errorMessage.setForeground(Color.RED);
        errorMessage.setBounds(50,250,200,20);
       
        employee = new JButton("Employee");
        employee.setBounds(270,240,120,25);
       
        f.add(employee);
        f.add(errorMessage);
        f.add(login);
        f.add(adminIdTextField);
        f.add(passwordField);
        f.add(titleLabel);
        f.add(adminIdLabel);
        f.add(passwordLabel);
       
        f.setSize(440,322);
        f.setLayout(null);
        f.setVisible(true);
       
        employee.addActionListener(new ActionListener(){
            //@Override
            public void actionPerformed(ActionEvent e) {
                f.setVisible(false);
                final JFrame eFrame = new JFrame();
                
                ImageIcon icon= new ImageIcon("images.png");
                JLabel pic= new JLabel();
                pic.setBounds(0,0,480,480);
                pic.setVisible(true);
                pic.setIcon(icon);
                eFrame.add(pic);
                
                
                final JLabel error= new JLabel("");
                error.setBounds(100,50,200,25);
                pic.add(error);
                
                final JLabel title = new JLabel("Employee Attendence");
                title.setBounds(100,100,200,25);
                pic.add(title);
                
                final JTextField text = new JTextField();
                text.setBounds(100,190,200,25);
                pic.add(text);
                
                final JButton select = new JButton("Select");
                select.setBounds(350,190,90,25);
                pic.add(select);
                
                final JButton submit = new JButton("Submit");
                submit.setBounds(120,230,90,25);
                pic.add(submit);
                
                final JButton logout = new JButton("Logout");
                logout.setBounds(230,230,90,25);
                pic.add(logout);
                
                eFrame.setLayout(null);
                eFrame.setSize(500,400);
                eFrame.setVisible(true);
                
                logout.addActionListener(new ActionListener()
                {
                    @Override
                    public void actionPerformed(ActionEvent e) 
                    {
                        eFrame.setVisible(false);
                        f.setVisible(true);
                    }
                    
                });
                
                select.addActionListener(new ActionListener(){
                    @Override
                    public void actionPerformed(ActionEvent e) 
                    {
                        error.setText("");  
                        final JFileChooser fileChooser = new JFileChooser();
                        int result = fileChooser.showOpenDialog(null);
                        if(result==JFileChooser.APPROVE_OPTION)
                        {
                            File selectedFile = fileChooser.getSelectedFile();
                            text.setText(selectedFile.getAbsolutePath());
                        }
                    }
                    
                });
                
                submit.addActionListener(new ActionListener()
                {
                    @Override
                    public void actionPerformed(ActionEvent e) 
                    {
                        if("".equals(text.getText())){
                            error.setForeground(Color.red);
                        error.setText("Please select some file!!");
                        }
                        else{
                          error.setText("");  
                        String filePath = text.getText();
                        String charset = "UTF-8";
                        Map hintMap = new HashMap();
                        hintMap.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.L);
                        try 
                        {
                            String empId = QRCode.readQRCode(filePath, charset, hintMap);
                            DateFormat df = new SimpleDateFormat("dd-MM-yyyy");
                            String todayDate = df.format(new Date());
                            
                            FileInputStream fin = new FileInputStream(todayDate + ".xlsx");                          
                            XSSFWorkbook wb = new XSSFWorkbook(fin);
                            XSSFSheet sheet = wb.getSheetAt(0);
                            int rowcount,colcount,i;
                            Row row;  
                            Cell cell;
                            String eid;
                            rowcount = sheet.getLastRowNum();
                            
                            final JFrame aFrame = new JFrame();
                            aFrame.setSize(300,300);
                            aFrame.setTitle("Entry frame");
                            
                            JLabel pic = new JLabel();
                            ImageIcon regicon= new ImageIcon("aframe.jpg");
                            pic.setIcon(regicon);
                            pic.setVisible(true);
                            pic.setBounds(0,0,300,300);
                            aFrame.add(pic);
                                    
                            final JLabel label = new JLabel();
                            label.setBounds(100,60,150,25);
                            label.setForeground(Color.red);
                            pic.add(label);
                                    
                            final JButton ok = new JButton("OK");
                            ok.setBounds(120,120,60,25);
                            pic.add(ok);
                            ok.addActionListener(new ActionListener()
                            {
                                @Override
                                public void actionPerformed(ActionEvent e) 
                                {
                                    aFrame.setVisible(false);
                                    text.setText("");
                                }
                                        
                            });
                            
                            for(i=1;i<=rowcount;i++)
                            {
                                row = sheet.getRow(i);
                                cell = row.getCell(0);
                                eid = cell.getStringCellValue();
                                if(eid.equals(empId))
                                {
                                    cell = row.getCell(6);
                                    String cellInput = cell.getStringCellValue();
                                    if(cellInput.equals("Absent"))
                                    {
                                        label.setText("Success entry") ;
                                        colcount = 5;
                                       
                                        Date date = new Date();
                                        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
                                        String time = sdf.format(date);
                                        cell = row.getCell(++colcount);
                                        cell.setCellValue((String)time);
                                        FileOutputStream fout = new FileOutputStream(todayDate + ".xlsx");
                                        wb.write(fout);
                                        
                                    }
                                    else if(row.getCell(7).getStringCellValue().equals("Absent"))
                                    {
                                        
                                        label.setBackground(Color.GREEN);
                                        label.setText("Thank you!");
                                        colcount = 7;
                                        Date date = new Date();
                                        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
                                        String time = sdf.format(date);
                                        cell = row.getCell(colcount);
                                        cell.setCellValue((String)time);
                                        FileOutputStream fout = new FileOutputStream(todayDate + ".xlsx");
                                        wb.write(fout);
                                        
                                    }
                                    else
                                    {
                                        label.setBackground(Color.RED);
                                        label.setText("You hava already exited");
                                    }
                                    
                                    break;
                                }
                            }
                            aFrame.setLayout(null);
                            aFrame.setSize(300,300);
                            aFrame.setVisible(true);
                        }
                       catch (IOException | NotFoundException ex) 
                        {
                            Logger.getLogger(EmployeeAttendance.class.getName()).log(Level.SEVERE, null, ex);
                        }
                       
                    }
                    
                }});
            }
           
       });
       
        login.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) 
            {
                System.out.println("In the login");
                String inputUserId,inputPassword;
                inputUserId = adminIdTextField.getText();
                inputPassword = passwordField.getText();
                if(userId.equals(inputUserId)&&password.equals(inputPassword))
                {
                    f.setVisible(false);
                    
                    final JFrame adminFrame = new JFrame();
                    adminFrame.setSize(1300,650);
                    adminFrame.setLayout(null);
                    adminFrame.setVisible(true);
                    
                    JPanel p1 = new JPanel();
                    p1.setBounds(5,5,400,640);
                    Color c1 = new Color(18,15,49);
                    p1.setBackground(c1);
                    p1.setLayout(null);
                    
           
                    final JLabel error= new JLabel("");
                    error.setBounds(60, 130, 250, 20);
                    error.setForeground(Color.red);
                    p1.add(error);
                  
                    JLabel regpic= new JLabel();
                    ImageIcon regicon= new ImageIcon("register.jpg");
                    regpic.setIcon(regicon);
                    regpic.setVisible(true);
                    regpic.setBounds(80,5,250,85);
                    p1.add(regpic);
                    
                    JLabel regpic1= new JLabel();
                    ImageIcon regicon1= new ImageIcon("1.jpg");
                    regpic1.setIcon(regicon1);
                    regpic1.setVisible(true);
                    regpic1.setBounds(0,170,400,320);
                    p1.add(regpic1); 
                    
                    JLabel p1Heading = new JLabel("NEW EMPLOYEE REGISTRATION");
                    p1Heading.setBounds(100,100,250,20);
                    p1Heading.setForeground(Color.white);
                    p1.add(p1Heading);
                    
                    final JLabel empId = new JLabel("EMP ID");
                    empId.setForeground(Color.white);
                    empId.setBounds(55,10,100,20);
                    regpic1.add(empId);
                    
                    final JTextField empIdTextField = new JTextField();
                    empIdTextField.setBounds(230,10,100,20);
                    regpic1.add(empIdTextField);
                   
                    JLabel firstName = new JLabel("FIRST NAME");
                    firstName.setBounds(55,50,100,20);
                    firstName.setForeground(Color.white);
                    regpic1.add(firstName);
                    final  JTextField firstNameTextField = new JTextField();
                    firstNameTextField.setBounds(230,50,100,20);
                    regpic1.add(firstNameTextField);
                   
                    JLabel lastName = new JLabel("LAST NAME");
                    lastName.setBounds(55,90,100,20);
                    lastName.setForeground(Color.white);
                    regpic1.add(lastName);
                    final JTextField lastNameTextField = new JTextField();
                    lastNameTextField.setBounds(230,90,100,20);
                    regpic1.add(lastNameTextField);
                   
                    JLabel email = new JLabel("EMAIL");
                    email.setBounds(55,130,100,20);
                    email.setForeground(Color.white);
                    regpic1.add(email);
                    final JTextField emailTextField = new JTextField();
                    emailTextField.setBounds(230,130,100,20);
                    regpic1.add(emailTextField);
                                     
                    JLabel phone = new JLabel("PHONE");
                    phone.setBounds(55,170,100,20);
                    phone.setForeground(Color.white);
                    regpic1.add(phone);
                    final JTextField phoneTextField = new JTextField();
                    phoneTextField.setBounds(230,170,100,20);
                    regpic1.add(phoneTextField);
                    
                   
                    JLabel designation = new JLabel("DESIGNATION");
                    designation.setBounds(55,210,100,20);
                    designation.setForeground(Color.white);
                    regpic1.add(designation);
                    final JTextField designationTextField = new JTextField();
                    designationTextField.setBounds(230,210,100,20);
                    regpic1.add(designationTextField);
                   
                    JButton submit = new JButton("SUBMIT");
                    submit.setBounds(55,260,100,30);
                    regpic1.add(submit);

                    JButton clear = new JButton("CLEAR");
                    clear.setBounds(230,260,100,30);
                    regpic1.add(clear);
                   
                    final JLabel text = new JLabel("");
                    text.setBounds(100,520,200,20);
                    text.setForeground(Color.white);
                    p1.add(text);
                   
                    submit.addActionListener(new ActionListener()
                    {
                        @Override
                        public void actionPerformed(ActionEvent e) 
                        {
                            System.out.println("In the submit");
                            Employee emp = new Employee();
                            emp.empId = empIdTextField.getText();
                            emp.firstName = firstNameTextField.getText();
                            emp.lastName = lastNameTextField.getText();
                            emp.email = emailTextField.getText();
                            emp.phone = phoneTextField.getText();
                            emp.designation = designationTextField.getText();
                            
                            
                            if("".equals(emp.empId)||"".equals(emp.firstName)||"".equals(emp.lastName)||"".equals(emp.email)||"".equals(emp.phone)||"".equals(emp.designation)){
                            error.setText("Please fill in all the fields!!");
                            
                            }
                            else{
                                
                            if((!isValidEmailAddress(emp.email))|| !isValidPhoneNumber(emp.phone)){
                                
                                        if((!isValidEmailAddress(emp.email))&&(!isValidPhoneNumber(emp.phone))){
                                        error.setText("Enter valid email and phone number!!!!");
                                       
                                        }
                                        else if(!isValidEmailAddress(emp.email)){
                                            error.setText("Enter valid email!!!!");
                                        }
                                        else{
                                        error.setText("Enter valid phone number!!!!");
                                        } 
                          
                            }
                            else{
                                File employeeDataFile = new File("EmployeeData.xlsx");
                                if(employeeDataFile.exists())
                                {
                                    try
                                    {
                                    FileInputStream filecheck = new FileInputStream(employeeDataFile);
                                    XSSFWorkbook wbcheck = new XSSFWorkbook(filecheck);
                                    XSSFSheet stcheck = wbcheck.getSheetAt(0);
                                    int rowCountcheck = stcheck.getLastRowNum();
                                    String id;
                                    boolean p=false;
                                    String empid=empIdTextField.getText();
                                    for(int i=0;i<rowCountcheck;i++)
                                    {
                                         Row r = stcheck.getRow(i+1);
                                         String value;
                                         id=r.getCell(0).getStringCellValue();
                                         if(empid.equals(id))
                                         {
                                         error.setText(" You have already assigned this EMPId to someone else");
                                         p=true;
                                         break;
                                         }
                                   }
                                    if(!p){
                                    error.setText(" ");
                            String QRCodeData = emp.empId;
                            String filePath = emp.firstName + ".png" ;
                            String charset = "UTF-8";
                            Map hintMap = new HashMap();
                            hintMap.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.L);
                            try 
                            {
                                QRCode.createQRCode(QRCodeData, filePath, charset, hintMap, 200, 200);
                            }   
                            catch ( WriterException | IOException ex) 
                            {
                               Logger.getLogger(EmployeeAttendance.class.getName()).log(Level.SEVERE, null, ex);
                            }
                           
                           
                            int rowcount,colcount;
                            File file = new File("EmployeeData.xlsx");
                            if(file.exists())
                            {
                                try
                                {
                                    FileInputStream inp = new FileInputStream(file);
                                    XSSFWorkbook wb = new XSSFWorkbook(inp);
                                    XSSFSheet sheet = wb.getSheetAt(0);
                                    rowcount = sheet.getLastRowNum();
                                    Row row = sheet.createRow(++rowcount);
                                    colcount=-1;
                                    Cell cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.empId );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.firstName );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.lastName );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.email );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.phone );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.designation );
                                    cell = row.createCell(++colcount);
                              
                                    FileOutputStream fout = new FileOutputStream("EmployeeData.xlsx");
                                    wb.write(fout);

                                    }
                                catch(IOException ex)
                                {
                                    System.out.println("Error occured");
                                }

                            }
                            else
                            {
                                try
                                {
                                    XSSFWorkbook wb = new XSSFWorkbook();
                                    XSSFSheet sheet = wb.createSheet("EmployeeData");
                                    rowcount=-1;
                                    colcount=-1;
                                    Row row = sheet.createRow(++rowcount);
                                    Cell cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "EMPLOYEE ID" );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "FIRST NAME" );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "LAST NAME" );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "EMAIL" );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "PHONE" );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "DESIGNATION" );
                                    cell = row.createCell(++colcount);
                                    row = sheet.createRow(++rowcount);
                                    colcount=-1;
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.empId );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.firstName );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.lastName );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.email );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.phone );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.designation );
                                    cell = row.createCell(++colcount);
                             

                                    FileOutputStream fout = new FileOutputStream("EmployeeData.xlsx");
                                    wb.write(fout);
                                }
                                catch(IOException er)
                                {
                                    System.out.println("Error occured");
                                }
                            }
                           
                            DateFormat df = new SimpleDateFormat("dd-MM-yyyy");
                            String todayDate = df.format(new Date());
                                  
                            File attendanceFile = new File(todayDate + ".xlsx");
                            if(!attendanceFile.exists())
                            {
                                XSSFWorkbook wb = new XSSFWorkbook();
                                XSSFSheet sheet = wb.createSheet("EmployeeAttendance");
                                rowcount=-1;
                                colcount=-1;
                                Row row = sheet.createRow(++rowcount);
                                Cell cell = row.createCell(++colcount);
                                cell.setCellValue((String) "EMPLOYEE ID" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "FIRST NAME" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "LAST NAME" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "EMAIL" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "PHONE" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "DESIGNATION" );
                                cell = row.createCell(++colcount);

                                cell.setCellValue((String) "TIME ENTRY" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "TIME EXIT" );
                               
                                try 
                                {
                                    FileOutputStream fout;
                                    fout = new FileOutputStream(todayDate + ".xlsx");
                                    wb.write(fout);
                                } 
                                catch (Exception ex) 
                                {
                                    Logger.getLogger(EmployeeAttendance.class.getName()).log(Level.SEVERE, null, ex);
                                }
                               
                            }
                           
                                  
                            try 
                            {
                                FileInputStream inp;
                                inp = new FileInputStream(attendanceFile);
                                XSSFWorkbook wb = new XSSFWorkbook(inp);

                                XSSFSheet sheet = wb.getSheetAt(0);
                                rowcount = sheet.getLastRowNum();
                                Row row = sheet.createRow(++rowcount);
                                colcount=-1;
                                Cell cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.empId );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.firstName );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.lastName );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.email );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.phone );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.designation );
                                cell = row.createCell(++colcount);

                                cell.setCellValue((String) "Absent" );
                                cell=row.createCell(++colcount);
                                cell.setCellValue((String) "Absent" );

                                FileOutputStream fout;
                                fout = new FileOutputStream(todayDate +  ".xlsx");
                                wb.write(fout);
                            } 
                            catch (Exception ex) 
                            {
                                Logger.getLogger(EmployeeAttendance.class.getName()).log(Level.SEVERE, null, ex);
                            }
                                  
                           
                            text.setForeground(Color.GREEN);
                            text.setText("Registered Successfully");
                            empIdTextField.setText(null);
                            firstNameTextField.setText(null);
                            lastNameTextField.setText(null);
                            emailTextField.setText(null);
                            phoneTextField.setText(null);
                            designationTextField.setText(null); 
                                    }
                                    
                                   }
                                    catch(Exception ex){System.out.println("Error occured");}
                                                
                                }
                                else{
                                error.setText(" ");
                            String QRCodeData = emp.empId;
                            String filePath = emp.firstName + ".png" ;
                            String charset = "UTF-8";
                            Map hintMap = new HashMap();
                            hintMap.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.L);
                            try 
                            {
                                QRCode.createQRCode(QRCodeData, filePath, charset, hintMap, 200, 200);
                            }   
                            catch ( WriterException | IOException ex) 
                            {
                               Logger.getLogger(EmployeeAttendance.class.getName()).log(Level.SEVERE, null, ex);
                            }
                           
                           
                            int rowcount,colcount;
                            File file = new File("EmployeeData.xlsx");
                            if(file.exists())
                            {
                                try
                                {
                                    FileInputStream inp = new FileInputStream(file);
                                    XSSFWorkbook wb = new XSSFWorkbook(inp);
                                    XSSFSheet sheet = wb.getSheetAt(0);
                                    rowcount = sheet.getLastRowNum();
                                    Row row = sheet.createRow(++rowcount);
                                    colcount=-1;
                                    Cell cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.empId );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.firstName );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.lastName );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.email );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.phone );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.designation );
                                    cell = row.createCell(++colcount);
                              
                                    FileOutputStream fout = new FileOutputStream("EmployeeData.xlsx");
                                    wb.write(fout);

                                    }
                                catch(IOException ex)
                                {
                                    System.out.println("Error occured");
                                }

                            }
                            else
                            {
                                try
                                {
                                    XSSFWorkbook wb = new XSSFWorkbook();
                                    XSSFSheet sheet = wb.createSheet("EmployeeData");
                                    rowcount=-1;
                                    colcount=-1;
                                    Row row = sheet.createRow(++rowcount);
                                    Cell cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "EMPLOYEE ID" );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "FIRST NAME" );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "LAST NAME" );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "EMAIL" );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "PHONE" );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) "DESIGNATION" );
                                    cell = row.createCell(++colcount);
                             


                                    row = sheet.createRow(++rowcount);
                                    colcount=-1;
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.empId );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.firstName );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.lastName );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.email );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.phone );
                                    cell = row.createCell(++colcount);
                                    cell.setCellValue((String) emp.designation );
                                    cell = row.createCell(++colcount);
                             

                                    FileOutputStream fout = new FileOutputStream("EmployeeData.xlsx");
                                    wb.write(fout);
                                }
                                catch(IOException er)
                                {
                                    System.out.println("Error occured");
                                }
                            }
                           
                            DateFormat df = new SimpleDateFormat("dd-MM-yyyy");
                            String todayDate = df.format(new Date());
                                  
                            File attendanceFile = new File(todayDate + ".xlsx");
                            if(!attendanceFile.exists())
                            {
                                XSSFWorkbook wb = new XSSFWorkbook();
                                XSSFSheet sheet = wb.createSheet("EmployeeAttendance");
                                rowcount=-1;
                                colcount=-1;
                                Row row = sheet.createRow(++rowcount);
                                Cell cell = row.createCell(++colcount);
                                cell.setCellValue((String) "EMPLOYEE ID" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "FIRST NAME" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "LAST NAME" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "EMAIL" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "PHONE" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "DESIGNATION" );
                                cell = row.createCell(++colcount);

                                cell.setCellValue((String) "TIME ENTRY" );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) "TIME EXIT" );
                               
                                try 
                                {
                                    FileOutputStream fout;
                                    fout = new FileOutputStream(todayDate + ".xlsx");
                                    wb.write(fout);
                                } 
                                catch (Exception ex) 
                                {
                                    Logger.getLogger(EmployeeAttendance.class.getName()).log(Level.SEVERE, null, ex);
                                }
                               
                            }
                           
                                  
                            try 
                            {
                                FileInputStream inp;
                                inp = new FileInputStream(attendanceFile);
                                XSSFWorkbook wb = new XSSFWorkbook(inp);

                                XSSFSheet sheet = wb.getSheetAt(0);
                                rowcount = sheet.getLastRowNum();
                                Row row = sheet.createRow(++rowcount);
                                colcount=-1;
                                Cell cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.empId );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.firstName );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.lastName );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.email );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.phone );
                                cell = row.createCell(++colcount);
                                cell.setCellValue((String) emp.designation );
                                cell = row.createCell(++colcount);

                                cell.setCellValue((String) "Absent" );
                                cell=row.createCell(++colcount);
                                cell.setCellValue((String) "Absent" );

                                FileOutputStream fout;
                                fout = new FileOutputStream(todayDate +  ".xlsx");
                                wb.write(fout);
                            } 
                            catch (Exception ex) 
                            {
                                Logger.getLogger(EmployeeAttendance.class.getName()).log(Level.SEVERE, null, ex);
                            }
                                  
                           
                            text.setForeground(Color.GREEN);
                            text.setText("Registered Successfully");
                            empIdTextField.setText(null);
                            firstNameTextField.setText(null);
                            lastNameTextField.setText(null);
                            emailTextField.setText(null);
                            phoneTextField.setText(null);
                            designationTextField.setText(null); /**/
                            }
                        }
                        
                        }
                            }});
                   
                    clear.addActionListener(new ActionListener(){
                        @Override
                        public void actionPerformed(ActionEvent e) {
                            empIdTextField.setText(null);
                            firstNameTextField.setText(null);
                            lastNameTextField.setText(null);
                            emailTextField.setText(null);
                            phoneTextField.setText(null);
                            designationTextField.setText(null);
                        }
                    });
                 
               
                   
                   
                    final JPanel p2 = new JPanel();
                    p2.setBounds(410,5,900,640);
                    p2.setLayout(null);
                   
                    
                     ImageIcon icon6= new ImageIcon("admin.png");
                    JLabel pic6 = new JLabel();
                    pic6.setBounds(0,0,900,640);
                    pic6.setVisible(true);
                    pic6.setIcon(icon6);
                    p2.add(pic6);
                    
                    ImageIcon icon= new ImageIcon("xy.png");
                    JLabel pic = new JLabel();
                    pic.setBounds(270,10,900,640);
                    pic.setVisible(true);
                    pic.setIcon(icon);
                    pic6.add(pic);
                    
                    JButton getattendance = new JButton("GET ATTENDANCE");
                    getattendance.setBounds(300,150,200,50);
                    pic6.add(getattendance);

                    JButton logout = new JButton("Logout");
                    logout.setBounds(600,100,100,25);
                    pic6.add(logout);
                    
                    
                    Color c = new Color(18,30,49);
                    p2.setBackground(c);
                   
                    logout.addActionListener(new ActionListener()
                    {
                        @Override
                        public void actionPerformed(ActionEvent e) {
                            adminFrame.setVisible(false);
                            f.setVisible(true);
                            adminIdTextField.setText(null);
                            passwordField.setText(null);
                        }
                       
                    });
                                 
                   
                    final JPanel p3 = new JPanel();
                    p3.setBounds(410,5,900,650);
                    p3.setLayout(null);
                    
                    ImageIcon icon7= new ImageIcon("admin.png");
                    JLabel pic7 = new JLabel();
                    pic7.setBounds(0,0,900,640);
                    pic7.setVisible(true);
                    pic7.setIcon(icon7);
                    p3.add(pic7);
                   
                    ImageIcon icon8= new ImageIcon("xy.png");
                    JLabel pic8= new JLabel();
                    pic8.setBounds(250,2,250,100);
                    pic8.setVisible(true);
                    pic8.setIcon(icon8);
                    pic7.add(pic8);

                    ImageIcon icon2= new ImageIcon("searchicon.png");
                    JLabel pic2= new JLabel();
                    pic2.setBounds(502,120,100,100);
                    pic2.setVisible(true);
                    pic2.setIcon(icon2);
                    pic7.add(pic2);
   
                    JLabel searchempattendance= new JLabel("Enter Emp id");
                    searchempattendance.setForeground(Color.WHITE);
                    searchempattendance.setBounds(150,150,100,25);
                    pic7.add(searchempattendance);

                    final JTextField searchtextfield= new JTextField();
                    searchtextfield.setBounds(250,150,100,25);
                    pic7.add(searchtextfield);

                    JButton search= new JButton("SEARCH");
                    search.setBounds(400,150,100,25);
                    pic7.add(search);

                    JButton logout1 = new JButton("Logout");
                    logout1.setBounds(600,100,100,25);
                    pic7.add(logout1);

                    final JLabel searchresult= new JLabel("");
                    searchresult.setBounds(150,15,500,200);
                    pic7.add(searchresult);

                    logout1.addActionListener(new ActionListener()
                    {
                        @Override
                        public void actionPerformed(ActionEvent e) {
                            adminFrame.setVisible(false);
                            f.setVisible(true);
                            adminIdTextField.setText(null);
                            passwordField.setText(null);
                        }
                       
                    });
                   
                    getattendance.addActionListener(new ActionListener()
                    {
                        @Override
                        public void actionPerformed(ActionEvent e)
                        {
                            p3.setVisible(true);
                            p2.setVisible(false);
                        }
                    }); 
                   
                    String[] columnNames = {"First Name", "Last Name","Email","Contact","Designation" ,"TimeEntry","TimeExit"};
                    String empfName="",emplName="",empemail="", empphn="", empds="", emptimeentry="", emptimeexit="";
                    Object[][] data1 = {{empfName, emplName, empemail,empphn, empds, emptimeentry,emptimeexit}};
                    final DefaultTableModel model = new DefaultTableModel(data1, columnNames);
                    final JTable table;
                    table= new JTable(model);
                    table.setForeground(Color.RED);
                    table.setBackground(Color.BLACK);
                    table.setFillsViewportHeight(true);
                    final JScrollPane scroll= new JScrollPane(table);
                    scroll.setBounds( 37, 200, 800, 350 ); 
                    pic7.add(scroll);
                    scroll.setVisible(false);
                    
                    search.addActionListener(new ActionListener()
                    {
                        @Override
                        public void actionPerformed(ActionEvent e)
                        {
                            String searchempid=searchtextfield.getText();
                    
                            try 
                            {      DateFormat df = new SimpleDateFormat("dd-MM-yyyy");
                                   String todayDate = df.format(new Date());
                                FileInputStream fin = new FileInputStream(todayDate + ".xlsx");                          
                                XSSFWorkbook wb = new XSSFWorkbook(fin);
                                XSSFSheet sheet = wb.getSheetAt(0);
                                int rowcount,colcount,i;
                                Row row;
                                Cell cell;
                                String eid , message;
                                rowcount = sheet.getLastRowNum();
                                if(searchempid.length()==0)
                                {
                                    searchresult.setForeground(Color.RED);
                                    searchresult.setText("Please enter the emp id!!");
                                }
                                else
                                {
                                    for(i=1;i<=rowcount;i++)
                                    {
                                        row = sheet.getRow(i);
                                        cell = row.getCell(0);
                                        eid = cell.getStringCellValue();

                                        if(eid.equals(searchempid))                                
                                        {   searchresult.setForeground(Color.GREEN);
                                            searchresult.setText("Your search results are as displayed! ");

                                            String empfName1,emplName1, empemail1,empphn1,empds1,empdate1,emptimeentry1,emptimeexit1;
                                            cell=row.getCell(1);
                                            empfName1=cell.getStringCellValue();

                                            cell=row.getCell(2);
                                            emplName1=cell.getStringCellValue();

                                            cell=row.getCell(3);
                                            empemail1=cell.getStringCellValue();

                                            cell=row.getCell(4);
                                            empphn1=cell.getStringCellValue();

                                            cell=row.getCell(5);
                                            empds1=cell.getStringCellValue();


                                            cell=row.getCell(6);
                                            emptimeentry1=cell.getStringCellValue();

                                            cell=row.getCell(7);
                                            emptimeexit1=cell.getStringCellValue();

                                            Object[] newRecord ={empfName1, emplName1, empemail1,empphn1, empds1, emptimeentry1,emptimeexit1};
                                            model.addRow(newRecord);

                                            scroll.setVisible(true);
                                            p3.add(scroll);
                                            break;

                                        }
                                    
                                    }
                                    if(i==(rowcount+1))
                                    {
                                        searchresult.setForeground(Color.RED);
                                        searchresult.setText("Sorry no such employee is registered!");
                                    }
                                }

                            }
                    
                            catch (IOException ex) 
                            {
                                Logger.getLogger(EmployeeAttendance.class.getName()).log(Level.SEVERE, null, ex);
                            }
                  
             
                        }
                    });
                   
                   
                    adminFrame.getContentPane().setBackground( c);
                    adminFrame.setLayout(null);
                    adminFrame.add(p1);
                    adminFrame.add(p2);
                    adminFrame.add(p3);
                   
                }
                else
                {
                    errorMessage.setText("Sorry try again");
                }
            }
       });
    }  
}
