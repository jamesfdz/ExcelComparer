package compare.excel;

import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.IOException;

import javax.swing.AbstractAction;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.UIManager;

public class MainActivity {
	
	static File[] excelSheetsFilePath;
	

	public static void main(String[] args) {
		//Setting up UI as per OS
		try {
	        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
	    }catch(Exception ex) {
	        ex.printStackTrace();
	    }
		
		final JFrame mainFrame = new JFrame();
		mainFrame.setSize(200, 70); 
		mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		JPanel panel = new JPanel();
		mainFrame.add(panel);
		mainFrame.setLayout(new GridLayout(0, 1));
		
		panel.add(new JButton(new AbstractAction("Browse") {

			@Override
			public void actionPerformed(ActionEvent arg0) {
				//Browse Functionality
				JFileChooser chooseFiles =new JFileChooser();
				chooseFiles.setMultiSelectionEnabled(true);
				int r = chooseFiles.showOpenDialog(new JFrame());
					if(r == chooseFiles.APPROVE_OPTION){
						excelSheetsFilePath = chooseFiles.getSelectedFiles();
					    int lengthOfFiles = excelSheetsFilePath.length;
					    if(lengthOfFiles != 0) {
					    	runComparer comparer = new runComparer();
							try {
								comparer.compare(excelSheetsFilePath);
							} catch (IOException e) {
								e.printStackTrace();
							}
							mainFrame.dispose();
					    }
					}
			}
        }));
		
		mainFrame.setLocationRelativeTo(null);
		mainFrame.setVisible(true);
	}
}
