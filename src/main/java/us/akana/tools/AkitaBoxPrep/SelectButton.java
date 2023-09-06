package us.akana.tools.AkitaBoxPrep;


import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.concurrent.ExecutionException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.SwingWorker;

/**
 * A file select button using swing - opens a file selection when clicked,
 * changes its text to reflect the selected file
 * 
 * @author Jaden Unruh
 */
@SuppressWarnings("serial")
class SelectButton extends JButton {
	SelectButton(final int whichSelect) {
		super(Messages.getString("SelectButton.Text")); //$NON-NLS-1$
		this.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				SwingWorker<Boolean, Void> sw = new SwingWorker<Boolean, Void>() {
					@Override
					protected Boolean doInBackground() throws Exception {
						JFileChooser fc = new JFileChooser();
						int returnVal = fc.showOpenDialog(Main.options);
						if (returnVal == JFileChooser.APPROVE_OPTION) {
							Main.selectedFiles[whichSelect] = fc.getSelectedFile();
							rename(fc.getSelectedFile().getName());
						}
						deSelected();
						return null;
					}
					
					@Override
					protected void done() {
						try {
							get();
						} catch (InterruptedException e) {
							e.printStackTrace();
						} catch (ExecutionException e) {
							e.getCause().printStackTrace();
							String[] choices = { Messages.getString("Main.Window.Error.Close"), //$NON-NLS-1$
									Messages.getString("Main.Window.Error.More") }; //$NON-NLS-1$
							Main.infoText = InfoText.ERROR;
							Main.info.setText(Main.getInfoText());
							if (JOptionPane.showOptionDialog(Main.options,
									String.format(Messages.getString("Main.Window.Error.ProblemLabel"), //$NON-NLS-1$
											e.getCause().toString()),
									Messages.getString("Main.Window.Error.Error"), JOptionPane.DEFAULT_OPTION, //$NON-NLS-1$
									JOptionPane.ERROR_MESSAGE, null, choices, choices[0]) == 1) {
								StringWriter sw = new StringWriter();
								e.printStackTrace(new PrintWriter(sw));
								JTextArea jta = new JTextArea(25, 50);
								jta.setText(String.format(Messages.getString("Main.Window.Error.FullTrace"), //$NON-NLS-1$
										sw.toString()));
								jta.setEditable(false);
								JOptionPane.showMessageDialog(Main.options, new JScrollPane(jta),
										Messages.getString("Main.Window.Error.Error"), JOptionPane.ERROR_MESSAGE); //$NON-NLS-1$
							}
						}
					}
				};
				selected();
				sw.execute();
			}
		});
	}

	/**
	 * Disables the button when selected
	 */
	void selected() {
		this.setEnabled(false);
	}

	/**
	 * Re-enables the button when deselected, and repacks Main.options to reflect
	 * the button's new width
	 */
	void deSelected() {
		this.setEnabled(true);
		Main.options.pack();
	}

	/**
	 * Renames the button to the given text
	 * 
	 * @param text the new name
	 */
	void rename(String text) {
		this.setText(text);
	}
}