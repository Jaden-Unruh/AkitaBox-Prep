package us.akana.tools.AkitaBoxPrep;

import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.concurrent.ExecutionException;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.SwingWorker;
import javax.swing.WindowConstants;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * The options for text in {@link Main#info}
 * 
 * @author Jaden Unruh
 */
enum InfoText {
	SELECT_PROMPT, ERROR, LOAD_SHEETS, ASSET_NUMS, DONE, SAVING
}

/**
 * Primary and entry class for AkitaBox-Prep, a tool to prepare data to be
 * imported into AkitaBox.
 * 
 * So far: Cross-Reference and update Inventory IDs with the Maximo ID from the
 * temporary IDs
 * 
 * @author Jaden Unruh
 */
public class Main {

	/**
	 * Formatter to pull a String from a cell in an XSSFSheet
	 */
	static final DataFormatter FORMATTER = new DataFormatter();

	/**
	 * Main program window
	 */
	static JFrame options;

	/**
	 * The Component Inventory and CA files, in that order
	 */
	static File[] selectedFiles = new File[2];

	/**
	 * XSSFWorkbook for Component Inventory and CA Files - loaded in
	 * {@link Main#loadSheets()}
	 * 
	 * @see Main#loadSheets()
	 */
	static XSSFWorkbook aspxBook, caBook;

	/**
	 * Information label at the bottom of {@link Main#options}
	 * 
	 * @see Main#options
	 * @see Main#infoText
	 */
	static JLabel info = new JLabel();

	/**
	 * The text currently showing in {@link Main#info}
	 */
	static InfoText infoText;

	/**
	 * Checks if the {@link Main#selectedFiles} are not null and are .xlsx files
	 * 
	 * @return true if both files are .xlsx
	 */
	static boolean checkCorrectSelections() {
		try {
			return isXLSX(selectedFiles[0]) && isXLSX(selectedFiles[1]);
		} catch (NullPointerException e) {
			return false;
		}
	}

	/**
	 * Pulls the temporary inventory id from each row of sheet two of
	 * {@link Main#caBook}, finds that id within {@link Main#aspxBook} and pulls the
	 * maximo id from that row of {@link Main#aspxBook}. Copies the maximo id back
	 * to column 3 of sheet 2 of {@link Main#caBook}. Writes comments to sheet 2 of
	 * {@link Main#caBook} if anything out of the ordinary happens.
	 */
	static void crossReferenceAssetNums() {
		updateInfo(InfoText.ASSET_NUMS);
		XSSFSheet caSheet = caBook.getSheetAt(1);
		for (int i = 1; i < caSheet.getPhysicalNumberOfRows(); i++) {
			XSSFRow activeRow = caSheet.getRow(i);
			String tempAssetId = FORMATTER.formatCellValue(activeRow.getCell(1));
			if (tempAssetId.endsWith("NEW")) { //$NON-NLS-1$
				String assetId = tempAssetId.substring(0, tempAssetId.length() - 3);
				String maximoId = findMaximoId(assetId);
				String oldVal = FORMATTER.formatCellValue(activeRow.getCell(2));
				if (!oldVal.equals(maximoId)) {
					if (oldVal.length() == 0)
						activeRow.getCell(2).setCellValue(maximoId);
					else if (maximoId.equals(Messages.getString("Main.Warn.AssetNotFound"))) //$NON-NLS-1$
						writeComment(caBook, caSheet, i, 2,
								String.format(Messages.getString("Main.Comment.AssetNotFound"), maximoId)); //$NON-NLS-1$
					else
						writeComment(caBook, caSheet, i, 2,
								String.format(Messages.getString("Main.Comment.InvalidMaximo"), maximoId, oldVal)); //$NON-NLS-1$
				}
			} else {
				// Write comment to tell user that items in column 2 of the CA file should end
				// in "NEW"
				writeComment(caBook, caSheet, i, 1, Messages.getString("Main.Comment.Ending")); //$NON-NLS-1$
			}
		}
	}

	/**
	 * Finds the maximo id associated with the temporary asset id within
	 * {@link Main#aspxBook}
	 * 
	 * @param assetId the assetId to use
	 * @return the associated maximo id
	 */
	static String findMaximoId(String assetId) {
		XSSFSheet invSheet = aspxBook.getSheetAt(0);
		for (int i = 1; i < invSheet.getPhysicalNumberOfRows(); i++) {
			XSSFRow activeRow = invSheet.getRow(i);
			if (FORMATTER.formatCellValue(activeRow.getCell(17)).equals(assetId))
				return FORMATTER.formatCellValue(activeRow.getCell(8));
		}
		return Messages.getString("Main.Warn.AssetNotFound"); //$NON-NLS-1$
	}

	/**
	 * The String associated with {@link Main#infoText}
	 * 
	 * @return the String currently showing in {@link Main#info}
	 * @see Main#infoText
	 */
	static String getInfoText() {
		switch (infoText) {
		case SELECT_PROMPT:
			return Messages.getString("Main.Info.SelectPrompt"); //$NON-NLS-1$
		case ERROR:
			return Messages.getString("Main.Info.Error"); //$NON-NLS-1$
		case LOAD_SHEETS:
			return Messages.getString("Main.Info.LoadSheets"); //$NON-NLS-1$
		case ASSET_NUMS:
			return Messages.getString("Main.Info.AssetNums"); //$NON-NLS-1$
		case DONE:
			return Messages.getString("Main.Info.Done"); //$NON-NLS-1$
		case SAVING:
			return Messages.getString("Main.Info.Saving"); //$NON-NLS-1$
		}
		return null;
	}

	/**
	 * Checks if the given file is of type XLSX (a microsoft excel workbook)
	 * 
	 * @param file the file to check
	 * @return true if the file is .xlsx
	 * @throws NullPointerException if the File is null
	 */
	static boolean isXLSX(File file) throws NullPointerException {
		return file.getName().toLowerCase().endsWith(".xlsx"); //$NON-NLS-1$
	}

	/**
	 * Loads {@link Main#aspxBook} and {@link Main#caBook} from their respective
	 * files - moving them to memory and making them readable in java
	 * 
	 * @throws IOException if reading data from the inputstream of either file fails
	 */
	static void loadSheets() throws IOException {
		updateInfo(InfoText.LOAD_SHEETS);
		aspxBook = new XSSFWorkbook(new FileInputStream(selectedFiles[0]));
		caBook = new XSSFWorkbook(new FileInputStream(selectedFiles[1]));
	}

	/**
	 * Entry method, calls {@link Main#openWindow()} to open {@link Main#options}
	 * 
	 * @param args unused
	 * @see Main#openWindow()
	 */
	public static void main(String[] args) {
		openWindow();
	}

	/**
	 * Adds contents to and opens {@link Main#options}
	 */
	private static void openWindow() {
		options = new JFrame(Messages.getString("Main.Window.Title")); //$NON-NLS-1$
		options.setSize(800, 700);
		options.setLayout(new GridBagLayout());
		options.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

		options.add(new JLabel(Messages.getString("Main.Window.ASPxPrompt")), //$NON-NLS-1$
				new GridBagConstraints(0, 0, 1, 1, 0, 0, GridBagConstraints.CENTER, 0, new Insets(0, 0, 0, 0), 0, 0));

		JButton selectASPx = new SelectButton(0);
		options.add(selectASPx,
				new GridBagConstraints(1, 0, 1, 1, 0, 0, GridBagConstraints.CENTER, 0, new Insets(0, 0, 0, 0), 0, 0));

		options.add(new JLabel(Messages.getString("Main.Window.CAPrompt")), //$NON-NLS-1$
				new GridBagConstraints(0, 1, 1, 1, 0, 0, GridBagConstraints.CENTER, 0, new Insets(0, 0, 0, 0), 0, 0));

		JButton selectCa = new SelectButton(1);
		options.add(selectCa,
				new GridBagConstraints(1, 1, 1, 1, 0, 0, GridBagConstraints.CENTER, 0, new Insets(0, 0, 0, 0), 0, 0));

		JButton cancel = new JButton(Messages.getString("Main.Window.Close")); //$NON-NLS-1$
		options.add(cancel,
				new GridBagConstraints(0, 4, 1, 1, 0, 0, GridBagConstraints.CENTER, 0, new Insets(0, 0, 0, 0), 0, 0));

		final JButton run = new JButton(Messages.getString("Main.Window.Open")); //$NON-NLS-1$
		options.add(run,
				new GridBagConstraints(1, 4, 1, 1, 0, 0, GridBagConstraints.CENTER, 0, new Insets(0, 0, 0, 0), 0, 0));

		options.add(info,
				new GridBagConstraints(0, 5, 2, 1, 0, 0, GridBagConstraints.CENTER, 0, new Insets(0, 0, 0, 0), 0, 0));

		cancel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.exit(0);
			}
		});

		run.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (checkCorrectSelections()) {
					SwingWorker<Boolean, Void> sw = new SwingWorker<Boolean, Void>() {
						@Override
						protected Boolean doInBackground() throws Exception {
							loadSheets();
							crossReferenceAssetNums();
							saveSheets();
							updateInfo(InfoText.DONE);
							run.setEnabled(true);
							return true;
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
								updateInfo(InfoText.ERROR);
								run.setEnabled(true);
								if (JOptionPane.showOptionDialog(options,
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
									JOptionPane.showMessageDialog(options, new JScrollPane(jta),
											Messages.getString("Main.Window.Error.Error"), JOptionPane.ERROR_MESSAGE); //$NON-NLS-1$
								}
							}
						}
					};
					run.setEnabled(false);
					sw.execute();
				} else {
					updateInfo(InfoText.SELECT_PROMPT);
				}
			}
		});

		options.pack();
		options.setVisible(true);
	}

	/**
	 * Saves {@link Main#caBook} with its updated content - copying from memory back
	 * to the disk
	 * 
	 * @throws FileNotFoundException if the user deleted the file between the time
	 *                               they selected it and its now being saved
	 * @throws IOException           if the write fails
	 */
	static void saveSheets() throws FileNotFoundException, IOException {
		updateInfo(InfoText.SAVING);
		FileOutputStream out = new FileOutputStream(selectedFiles[1]);
		caBook.write(out);
		out.close();
		caBook.close();
	}

	/**
	 * Updates the text of {@link Main#info}
	 * 
	 * @param text the new text
	 */
	static void updateInfo(InfoText text) {
		infoText = text;
		info.setText(getInfoText());
		options.pack();
	}

	/**
	 * Writes a comment to the given sheet
	 * 
	 * @param book    the parent book of the sheet to write to
	 * @param sheet   the sheet to write the comment to
	 * @param row     the row of the cell to write the comment to
	 * @param col     the column of the cell to write the comment to
	 * @param message the desired contents of the commment
	 */
	static void writeComment(XSSFWorkbook book, XSSFSheet sheet, int row, int col, String message) {
		CreationHelper factory = book.getCreationHelper();
		ClientAnchor anchor = factory.createClientAnchor();
		anchor.setCol1(col + 1);
		anchor.setCol2(col + 3);
		anchor.setRow1(row + 1);
		anchor.setRow2(row + 3);
		XSSFDrawing drawing = sheet.createDrawingPatriarch();
		Comment comment = drawing.createCellComment(anchor);
		comment.setString(factory.createRichTextString(message));
		comment.setAuthor(Messages.getString("Main.Comment.Author")); //$NON-NLS-1$
		sheet.getRow(row).getCell(col).setCellComment(comment);
	}
}