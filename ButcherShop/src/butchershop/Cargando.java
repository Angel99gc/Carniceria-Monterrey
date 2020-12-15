/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package butchershop;

import javax.swing.JLabel;

/**
 *
 * @author Angel
 */
public class Cargando implements Runnable{
    JLabel label;

    public Cargando(JLabel label) {
        this.label = label;
    }
    
    public void show(){
        new Thread(this).start();
    }
    @Override
    public void run() {
        
    }
    
}
