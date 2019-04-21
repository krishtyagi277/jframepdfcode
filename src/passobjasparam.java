/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Krishna Tyagi
 */
public class passobjasparam {
   
      static String arr[] ={"Now","all","aid","come"};
      int a=0;
    public static void main(String args[]){
   
     for(int i=0;i<arr.length;i++)
     {
        for(int j=i+1;j<arr.length;j++)
        {
        
        if(arr[j].compareTo(arr[i])<0)
        {
        String t =arr[i];
        arr[i] =arr[j];
        arr[j] =t;
        
        }
        }     
        ;
     System.out.println(arr[i]);
     }
         
        
        
    }
}
