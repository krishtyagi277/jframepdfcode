/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package jframepdfcode;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Krishna Tyagi
 */
public class PdfTest extends javax.swing.JFrame {

     static Object[] Data1 = {"Name","","","","","","","JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC","YEAR"};
   static Object[][] Data = new Object[1000][1000];
   static Double[][] count=new Double[1000][1000];
   static Double[][] count2=new Double[1000][1000];
   static Double[] amt=new Double[40];
   static Object[] nmonths = {"JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC","YEAR"}; 
   static Object[][] newData=new Object[100][100];
   static String[] Date1={"01","02","03","04","05","06","07","08","09","10","11","12"};
   static int[] Date2={1,2,3,4,5,6,7,8,9,10,11,12};
   static String[] year=new String[1000];
   static String filename;
    static String filepath;
   static String fieldvalue=null;
   
   
   public PdfTest() {
       super("pdf to excel"); 
       initComponents();
             
    }

  public void Convert_method(String filepth,String filenm) throws FileNotFoundException, IOException, InterruptedException{
    
        
     HashMap< String,Integer> hmap = new HashMap< String,Integer>();       
 XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Deposit Data");
      
        System.arraycopy(Data1, 0, Data[0], 0,Data1.length);

int ch=0,i=0,i2=0,i3=1,totalname=0,copyname=0,nocopy=0;
     
        try{
   
        PdfManager pdfManager = new PdfManager();
        
        //  String pdfFile="E:\\depositdoc_"+ch+".pdf"; "E:\\estatement JAN TO APRIL 2018-2.pdf"
     String pdfFile=filepth; 
     pdfManager.setFilePath( pdfFile);
     String docText = pdfManager.ToText();
      int temp=0;    
      int inc=19,monStart=7,monEnd=19,amtinc=1;
     int col=7,mcol=19,fDate=0;
     int e=0,d_sym,j=1;
     Double tm=0.0,totalamt=0.0;
     String depYear,depDate,Date;
     Double amount;
     String b="";
     int tmp=1,tmp2=0;
  
     while(temp==0){
  
       d_sym=docText.indexOf("Deposit",e);
         if(d_sym>=0){
             
            
            depDate=docText.substring(d_sym-7,d_sym);            
                 int tp=depDate.lastIndexOf("/");
                   if(tp!=-1)
                   {
                    depYear=depDate.substring(tp+1,tp+3);
                      Date=depDate.substring(tp-2,tp);
                       for(int dt=0;dt<12;dt++)
                       {
                        if(Date.equals(Date1[dt]))
                        {
                          fDate=Date2[dt];
                          break; 
                        }  
                       }
                         depYear="20"+depYear;
                         year[mcol]=depYear; 
                        
                
             
         int c=docText.indexOf("\n",d_sym);
            String d=docText.substring(d_sym,docText.indexOf(".",d_sym));
               String fname;
                   if(d.contains("ib."))
                    {
                     fname=d;
                    }
                   else
                   {
                        int sp=d.lastIndexOf(" ");
                         int deposit=d.indexOf("Deposit");
                        fname=d.substring(deposit+7,sp);
                   }
                    
         int nametmp=0;Double nameAmt=0.0;
          totalname++;
             if(hmap.containsKey(fname))
             {
              ch=hmap.get(fname);
               Data[ch][0]=fname;
                nametmp++;
                nameAmt=count[ch][mcol];
               copyname++;
             }
             else
             {
              i++;
               hmap.put(fname, i);
                  Data[i][0]=fname;
                 nocopy++;
                
            }
  
           
        
        String rupe,rupe2,depAmt;
        int last_sp,nex_dot;
           
                int dot=docText.indexOf(".",d_sym);
                  rupe=docText.substring(dot-9,dot+3);
                   if(rupe.contains("ib.")||rupe.contains("C."))
                      {
                         int dot2=docText.indexOf(".",dot+1);
                        rupe=docText.substring(dot,dot2+3);
                        last_sp=rupe.lastIndexOf(" ");
                        nex_dot=rupe.indexOf(".",last_sp);
                        rupe2=rupe.substring(last_sp,nex_dot+3);
                      }else{
                last_sp = rupe.lastIndexOf(" ");
                nex_dot=rupe.indexOf(".");
                rupe2=rupe.substring(last_sp,nex_dot+3);
                      
                     
                 }
   
  depAmt=rupe2.replaceAll("[^0-9.]","");        
            
 e=dot;

  
       
          if(depYear.equals(b)||tmp==1)
             {   
                 
                amtinc=1;
                totalamt=0.0;
                int col4=col;
                  if(nametmp!=0)
                  {
                    totalamt=nameAmt;
                      for(int month1=monStart;month1<monEnd;month1++)
                      {
                         if(count[ch][month1]!=null)
                         {
                           amt[amtinc]=count[ch][month1];
                         }
                         else{
                               amt[amtinc]=0.0;
                             }
                        amtinc++;
                      }
                    
                            for(int month=1;month<13;month++)
                            {
                             if(fDate==month)
                             { 
                                 amount=Double.parseDouble(depAmt);
                                   amt[month]=amt[month]+amount;
                                      Data[ch][col4]=amt[month];
                                   count[ch][col4]=amt[month];
                                 if(totalamt!=null){
                                   totalamt=totalamt+amt[month];
                                 }break;
                             }
                              else{
                               col4++;
                                  }
                            }
                 
           
                   }
                  else{
                       totalamt=0.0;
                          for(int month2=1;month2<13;month2++)
                          {
                            amt[month2]=0.0;
                          }
                            for(int month=1;month<13;month++)
                           {
                             if(fDate==month)
                             { 
                               amount=Double.parseDouble(depAmt);
                                amt[month]=amt[month]+amount;
                                 Data[i][col4]=amt[month];
                                   count[i][col4]=amt[month];
                                 totalamt=totalamt+amt[month];
                                break;
                             }
                            else{
                            col4++;
                             }
                           }          
                                     }
                       tmp++;
                       if(totalamt!=null){
                 Data[i][inc]=totalamt.toString();
                     count[i][inc]=totalamt;
                       }
                             Data[0][inc]=b;
           //    totalamt=totalamt+amount;               
               }
         else{ 
              Data[0][inc]=b;
                  monStart=monStart+13;
                    monEnd=monEnd+13;
                      inc=inc+13;
                          col=col+13;
                       int col2=col;
                      int col3=col;
                     mcol=mcol+13;
                    amtinc=1;
              for(int y=0;y<13;y++)
              {
              Data[0][col2]=nmonths[y];
              col2++;
              }
                 if(nametmp!=0){
                  totalamt=(Double)Data[ch][mcol];
                     for(int month1=monStart;month1<monEnd;month1++)
                    {
                       if(count[ch][month1]!=null)
                        {
                         amt[amtinc]=count[ch][month1];
                        }
                        else
                       {
                          amt[amtinc]=0.0;
                        }
                       amtinc++;
                    }
                  for(int month=1;month<13;month++)
                    {
                         if(fDate==month)
                         { 
                            amount=Double.parseDouble(depAmt);
                              amt[month]=amt[month]+amount;
                               Data[ch][col3]=amt[month];
                              count[ch][col3]=amt[month];
                          if(totalamt!=null){
                             totalamt=totalamt+amt[month];
                          }
                             break;
                         }
             
                         else
                         {
                          col3++;
                         }
                    }
              
                }
                  else{
                       totalamt=0.0;
                           for(int month2=1;month2<13;month2++)
                           {
                             amt[month2]=0.0;
                           }
                           for(int month=1;month<13;month++)
                           {  
                               if(fDate==month)
                              { 
                               amount=Double.parseDouble(depAmt);
                                amt[month]=amt[month]+amount;
                                 Data[i][col3]=amt[month];
                                  count[i][col3]=amt[month];
                                totalamt=totalamt+amt[month];
                               break;
                              }
                          else{
                             col3++;
                              }
                           }
                  Data[i][inc]=totalamt.toString();
                 count[i][inc]=totalamt;      
                 }
         
                
         
             }    
          
       
temp=0;
j++;
           
b=depYear;
     
         }else{tmp2++;
                int dot4=docText.indexOf(".",d_sym);
                String dep_wdt=docText.substring(d_sym,dot4-4);
                       String rupe3=docText.substring(dot4-8,dot4+3);
                       int l_sp=rupe3.lastIndexOf(" ");
                       int l_dot=rupe3.indexOf(".",l_sp);
                String dep_amt2=rupe3.substring(l_sp,l_dot+3);
                                      System.out.println(""+dep_amt2);
                dep_amt2=dep_amt2.replaceAll("[^0-9.]","");
                  while(i2<i3){
                  newData[i2][0]=dep_wdt;
                  newData[i2][19]=dep_amt2;
                  count2[i2][19]=Double.parseDouble(dep_amt2);
                  i2++;
                  }
                  i3++;
                  totalname++;
                temp=0;
               e=dot4;    
                   }
                      }
   else{
   temp=1;
    
    
       }
   }
          
     Double tamt;int z=7;
    for(int k=1;k<=totalname;k++){
        tamt=0.0;
        z=7; 
     while(z<=mcol){
       
      if(Data[0][z]!=null){
        String d=Data[0][z].toString();
        if(!d.contains("20")){
            if(count[k][z]!=null){
        tamt=tamt+count[k][z];        
           }      
        z++;
        }
        else{ 
            if(tamt!=0.0){
        Data[k][z]=tamt;
        count[k][z]=tamt;
            }
        tamt=0.0;
        z++;
        }
      }
     }
    }
    
     
     
     
   if(tmp2!=0){
     i2=0;
    i++;
    while(i2<i3){
    Data[i][0]=newData[i2][0];
    if(newData[i2][19]!=null){
    Data[i][19]=newData[i2][19];
    count[i][19]=count[i2][19];
    }
  i2++;
  i++;
            }
   }   
         
     
     
     
int rowCount = 0;
         
        for (Object[] aBook : Data) {
            XSSFRow row = sheet.createRow(rowCount);
            int columnCount = 0;
             
            for (Object field : aBook) {
             
                XSSFCell cell = row.createCell(columnCount);
              
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Double) {
                    if((Double) field!=0.0){
                    cell.setCellValue((Double) field);
                    }
                }
           
                columnCount++;
            }
            rowCount++; 
        }
        try( FileOutputStream output = new FileOutputStream (filenm+".xlsx")){
           
       workbook.write(output);
           
    } 
     
       Thread.sleep(10000);
      jLabel2.setText("file converted Successfully...........");
 
      
   } catch (NumberFormatException t)
        {
            t.printStackTrace();
            
        }
        catch(Exception e){
        JOptionPane.showInternalMessageDialog(null, e);
        }
    
    
    }
    
    
    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jTextField1 = new javax.swing.JTextField();
        choose_btn = new javax.swing.JButton();
        convert_btn = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jTextField2 = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setBackground(new java.awt.Color(51, 0, 204));

        jTextField1.setEditable(false);
        jTextField1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField1ActionPerformed(evt);
            }
        });

        choose_btn.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        choose_btn.setText("choose pdf");
        choose_btn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                choose_btnActionPerformed(evt);
            }
        });

        convert_btn.setFont(new java.awt.Font("Tahoma", 1, 14)); // NOI18N
        convert_btn.setText("convert");
        convert_btn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                convert_btnActionPerformed(evt);
            }
        });

        jLabel1.setFont(new java.awt.Font("Tahoma", 1, 14)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(204, 0, 0));
        jLabel1.setText("PDF TO EXCEL");

        jLabel2.setFont(new java.awt.Font("Tahoma", 1, 14)); // NOI18N
        jLabel2.setForeground(new java.awt.Color(51, 0, 153));
        jLabel2.setText("Choose your file and press the button ..");

        jTextField2.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        jTextField2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField2ActionPerformed(evt);
            }
        });

        jLabel3.setFont(new java.awt.Font("Tahoma", 1, 14)); // NOI18N
        jLabel3.setForeground(new java.awt.Color(51, 51, 255));
        jLabel3.setText("Enter your file name:");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(24, 24, 24)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 359, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, 329, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(choose_btn, javax.swing.GroupLayout.PREFERRED_SIZE, 107, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 146, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, 329, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(203, 203, 203)
                        .addComponent(convert_btn, javax.swing.GroupLayout.PREFERRED_SIZE, 87, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(189, 189, 189)
                        .addComponent(jLabel1)))
                .addContainerGap(26, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 34, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 46, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(choose_btn, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 19, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, 36, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(convert_btn, javax.swing.GroupLayout.PREFERRED_SIZE, 40, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 45, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void choose_btnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_choose_btnActionPerformed
     
          
        JFileChooser choose =new JFileChooser();
        choose.showOpenDialog(null);
        File f=choose.getSelectedFile();
        filepath=f.getAbsolutePath();
        jTextField1.setText(filepath);
        
    }//GEN-LAST:event_choose_btnActionPerformed

    private void jTextField1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField1ActionPerformed
  
    }//GEN-LAST:event_jTextField1ActionPerformed

    private void convert_btnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_convert_btnActionPerformed

         try { 
             jLabel2.setText("converting....wait");
             filename= jTextField2.getText();
            filename= filename.trim();
             Convert_method(filepath,filename);
       
         } catch (IOException ex) {
             Logger.getLogger(PdfTest.class.getName()).log(Level.SEVERE, null, ex);
            
        JOptionPane.showInternalMessageDialog(null, ex);
        
         } catch (InterruptedException ex) {
             Logger.getLogger(PdfTest.class.getName()).log(Level.SEVERE, null, ex);
         }
        
    }//GEN-LAST:event_convert_btnActionPerformed

    private void jTextField2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField2ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField2ActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(PdfTest.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(PdfTest.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(PdfTest.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(PdfTest.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new PdfTest().setVisible(true);
            }
        });
        
        
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton choose_btn;
    private javax.swing.JButton convert_btn;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTextField jTextField2;
    // End of variables declaration//GEN-END:variables
}
