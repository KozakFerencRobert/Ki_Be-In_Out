package Vendegek_Belepok;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import javax.swing.table.AbstractTableModel;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row;

public class Vendegek_Es_Belepok extends JFrame {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private Workbook workbook;
	private Sheet sheet;
	private int rowCount;
	private ExcelTableModel tableModel;

	public Vendegek_Es_Belepok() {
		super("Vendégek És Belépők Nyilvántartása");

		try {
			String currentDirectory = System.getProperty("user.dir");
			String folderName = "Vendégek_És_Belépők";
			String folderPath = currentDirectory + File.separator + folderName;
			String today = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
			File file = new File(folderPath + File.separator + "Vendég_Belépő_" + today + ".xlsx");

			Path folder = Paths.get(folderPath);
			if (!Files.exists(folder)) {
				Files.createDirectory(folder);
			}
			if (file.exists()) {
				FileInputStream fis = new FileInputStream(file);
				workbook = new XSSFWorkbook(fis);
				sheet = workbook.getSheetAt(0);
				fis.close();
				rowCount = sheet.getLastRowNum() + 1;
			} else {
				workbook = new XSSFWorkbook();
				sheet = workbook.createSheet("Adatok");

				rowCount = 0;
				Row nulla = sheet.createRow(rowCount++);

				// A cella stílusának beállítása a szegélyezéshez
				CellStyle cellStyle = workbook.createCellStyle();
				cellStyle.setBorderTop(BorderStyle.THIN);
				cellStyle.setBorderBottom(BorderStyle.THIN);
				cellStyle.setBorderLeft(BorderStyle.THIN);
				cellStyle.setBorderRight(BorderStyle.THIN);

				nulla.createCell(0).setCellValue("-------------------------");// céginfó helye
				sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));

				rowCount = 2;

				Row egy = sheet.createRow(rowCount);
				egy.setRowStyle(cellStyle);
				egy.createCell(0).setCellValue("VENDÉGEK ÉS /BELÉPŐK/ NYILVÁNTARTÁSA");
				sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 6));

				rowCount = 3;
				Row elso = sheet.createRow(rowCount++);
				elso.createCell(0).setCellValue("Sorszám:");
				elso.createCell(1).setCellValue("Belépett(év/hó/nap/óra/perc):");
				elso.createCell(2).setCellValue("Vendég neve:");
				elso.createCell(3).setCellValue("Kártya száma:");
				elso.createCell(4).setCellValue("Ügyintéző neve(akihez érkezett):");
				elso.createCell(5).setCellValue("Kilépett(nap/óra/perc):");
				elso.createCell(6).setCellValue("(Visitor) kártya visszavétele:");

				for (Cell cell : elso) {
					sheet.autoSizeColumn(cell.getColumnIndex());
					cell.setCellStyle(cellStyle);// Oszlop szélesség automatikus beállítása
				}

			}

			createUI();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// ---------------------------------------------------------------------------------------
	private String correctPassword = "abcd1234"; // Az itt megadott jelszó helyes jelszó

	private void kilepesButtonClicked() {
		int dialogResult = JOptionPane.showConfirmDialog(this, "Biztosan ki szeretne lépni?", "Kilépés megerősítése",
				JOptionPane.YES_NO_OPTION);
		if (dialogResult == JOptionPane.YES_OPTION) {
			// Ha a felhasználó igennel válaszolt, bezárjuk az alkalmazást
			dispose();
		}
	}

	private int getNextSorszam() {
		int maxSorszam = 0;
		for (Row row : sheet) {
			if (row.getRowNum() == 0)
				continue; // Fejléc sort kihagyjuk

			Cell sorszamCell = row.getCell(0);
			if (sorszamCell != null && sorszamCell.getCellType() == CellType.NUMERIC) {
				int cellValue = (int) sorszamCell.getNumericCellValue();
				if (cellValue > maxSorszam) {
					maxSorszam = cellValue;
				}
			}
		}
		return maxSorszam + 1;
	}

	private void showModifyDialog() {
		JTextField sorszamField = new JTextField(10);
		Object[] message = { "Adja meg a módosítani kívánt sorszámot:", sorszamField };
		int option = JOptionPane.showConfirmDialog(this, message, "Módosítás", JOptionPane.OK_CANCEL_OPTION);

		if (option == JOptionPane.OK_OPTION) {
			String sorszamString = sorszamField.getText();
			if (!sorszamString.isEmpty()) {
				try {
					int sorszam = Integer.parseInt(sorszamString);

					// Ellenőrizd a jelszót
					JPasswordField passwordField = new JPasswordField(20);
					int passwordOption = JOptionPane.showConfirmDialog(this,
							new Object[] { "Adja meg a jelszót:", passwordField }, "Jelszó ellenőrzése",
							JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);

					if (passwordOption == JOptionPane.OK_OPTION) {
						char[] enteredPasswordChars = passwordField.getPassword();
						String enteredPassword = new String(enteredPasswordChars);

						if (checkPassword(enteredPassword)) {
							// Végrehajtsd a módosítást
							modifyRow(sorszam);
						} else {
							JOptionPane.showMessageDialog(this, "Helytelen jelszó!", "Hiba", JOptionPane.ERROR_MESSAGE);
						}
					}
				} catch (NumberFormatException ex) {
					JOptionPane.showMessageDialog(this, "Érvénytelen sorszám formátum! Adjon meg egy egész számot.",
							"Hiba", JOptionPane.ERROR_MESSAGE);
				}
			} else {
				JOptionPane.showMessageDialog(this, "A sorszám mező nem lehet üres.", "Hiba",
						JOptionPane.ERROR_MESSAGE);
			}
		}
	}

	private boolean checkPassword(String enteredPassword) {
		return enteredPassword.equals(correctPassword);
	}

	private void modifyRow(int sorszam) {
		Row row = getRowBySorszam(sorszam);

		if (row != null) {
			JTextField beField = new JTextField(10);
			JTextField kiField = new JTextField(10);
			JTextField VendegField = new JTextField(20);
			JTextField KartyaField = new JTextField(20);
			JTextField UgyintezoField = new JTextField(10);
			JTextField VisitorField = new JTextField(30);

			Object[] message = { "Belépett(év/hó/nap/óra/perc):", beField, "Vendég neve:", VendegField, "Kártya száma:",
					KartyaField, "Ügyintéző neve(akihez érkezett):", UgyintezoField, "Kilépett(nap/óra/perc):", kiField,
					"(Visitor) kártya visszavétele:", VisitorField, };

			// A cella stílusának beállítása a szegélyezéshez
			CellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setBorderTop(BorderStyle.THIN);
			cellStyle.setBorderBottom(BorderStyle.THIN);
			cellStyle.setBorderLeft(BorderStyle.THIN);
			cellStyle.setBorderRight(BorderStyle.THIN);

			int option = JOptionPane.showConfirmDialog(this, message, "Sor módosítása", JOptionPane.OK_CANCEL_OPTION);
			int dialogResult = JOptionPane.showConfirmDialog(null,
					"Biztosan módosítani akarja a sort? \n A módosítás nem visszavonható!!", "Törlés megerősítése",
					JOptionPane.YES_NO_OPTION);

			if (dialogResult == JOptionPane.YES_OPTION) {
				if (option == JOptionPane.OK_OPTION) {
					String belepes = beField.getText();
					String vendegnev = VendegField.getText();
					String kilepes = kiField.getText();
					String kartya = KartyaField.getText();
					String ugyintezo = UgyintezoField.getText();
					String visitor = VisitorField.getText();
					// Check if all fields are filled
					if (belepes.isEmpty() && vendegnev.isEmpty() && kilepes.isEmpty() && kartya.isEmpty()
							&& ugyintezo.isEmpty() && visitor.isEmpty()) {
						JOptionPane.showMessageDialog(this, "Minden mezőt ki kell tölteni a módosításhoz!", "Hiba",
								JOptionPane.ERROR_MESSAGE);
						return;
					}

					row.getCell(1).setCellValue(belepes);
					row.getCell(2).setCellValue(vendegnev);
					row.getCell(3).setCellValue(kartya);
					row.getCell(4).setCellValue(ugyintezo);
					row.getCell(5).setCellValue(kilepes);
					row.getCell(6).setCellValue(visitor);

					for (Cell cell : row) {
						cell.setCellStyle(cellStyle); // Az új cella stílus hozzáadása
					}

					for (Cell cell : row) {
						sheet.autoSizeColumn(cell.getColumnIndex()); // Oszlop szélesség automatikus beállítása
					}
					saveExcelFile();

					JOptionPane.showMessageDialog(this, "Sor módosítva.", "Siker", JOptionPane.INFORMATION_MESSAGE);
				}
			}
		} else {
			JOptionPane.showMessageDialog(this, "A megadott sorszám nem található az Excel táblázatban.", "Hiba",
					JOptionPane.ERROR_MESSAGE);
		}
	}

	private void createUI() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLayout(new BorderLayout());

		// A cella stílusának beállítása a szegélyezéshez
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);

		tableModel = new ExcelTableModel(sheet);
		JTable table = new JTable(tableModel);

		table.setBackground(new Color(190,195,198));

		
		// A táblázat hozzáadása görgethető panelhez
		JScrollPane scrollPane = new JScrollPane(table);
		add(scrollPane, BorderLayout.CENTER);

		// Beviteli mezők létrehozása és hozzáadása a panelhez
		JTextField sorszamField = new JTextField(10);
		JTextField beField = new JTextField(10);
		JTextField vendegField = new JTextField(20);
		JTextField kartyaField = new JTextField(10);
		JTextField ugyintezoField = new JTextField(20);
		JTextField kiField = new JTextField(20);
		JTextField visitorField = new JTextField(10);

		JButton hozzaadButton = new JButton("Hozzáad");
		hozzaadButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				String belepes = beField.getText();
				String vendegnev = vendegField.getText();
				String kilepes = kiField.getText();
				String kartya = kartyaField.getText();
				String ugyintezo = ugyintezoField.getText();
				String visitor = visitorField.getText();

				int sorszam = getNextSorszam();
				sorszamField.setText(String.valueOf(sorszam));

				// Ellenőrizzük, hogy minden mező kitöltve van-e
				if (!belepes.isEmpty() || !kartya.isEmpty() || !ugyintezo.isEmpty() || !vendegnev.isEmpty()) {

					try {
						// int sorszamInt = Integer.parseInt(sorszam);

						// Ellenőrizzük, hogy van-e már adott sorszámú sor az Excel táblázatban
						Row row = getRowBySorszam(sorszam);

						if (row == null) {
							// Ha nincs, akkor létrehozzuk egy új sort és beírjuk az adatokat
							row = sheet.createRow(rowCount++);
							row.createCell(0).setCellValue(sorszam);
							row.createCell(1).setCellValue(belepes);
							row.createCell(2).setCellValue(vendegnev);
							row.createCell(3).setCellValue(kartya);
							row.createCell(4).setCellValue(ugyintezo);
							row.createCell(5).setCellValue(kilepes);
							row.createCell(6).setCellValue(visitor);

							for (Cell cell : row) {
								sheet.autoSizeColumn(cell.getColumnIndex());
								cell.setCellStyle(cellStyle);// Oszlop szélesség automatikus beállítása
							}
							// A változtatások mentése az Excel fájlba
							saveExcelFile();

							// Frissítjük a táblázatot az új adatokkal
							table.setModel(new ExcelTableModel(sheet));
							table.repaint();
							// Mezők tartalmának törlése az adatbevitel után
							sorszamField.setText("");
							beField.setText("");
							kiField.setText("");
							vendegField.setText("");
							kartyaField.setText("");
							ugyintezoField.setText("");
							visitorField.setText("");

						}
					} catch (NumberFormatException ex) {
						JOptionPane.showMessageDialog(Vendegek_Es_Belepok.this, "A Sorszám csak szám lehet!");
					}
				} else {
					JOptionPane.showMessageDialog(Vendegek_Es_Belepok.this, "Minden mezőt ki kell tölteni!");
				}
			}
		});

		JButton modositasButton = new JButton("Módosítás");
		modositasButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				showModifyDialog();
			}
		});

		JButton torlesButton = new JButton("Törlés");
		torlesButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				JPasswordField passwordField = new JPasswordField(); // Használd JPasswordField-t
				int passwordResult = JOptionPane.showConfirmDialog(null,
						new Object[] { "Adja meg a jelszót:", passwordField }, "Jelszó ellenőrzése",
						JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
				if (passwordResult == JOptionPane.OK_OPTION) {
					char[] enteredPasswordChars = passwordField.getPassword(); // Jelszó karaktertömb lekérése
					String enteredPassword = new String(enteredPasswordChars); // Karaktertömb átalakítása Stringgé
					Arrays.fill(enteredPasswordChars, '\0'); // Törlés a karaktertömbből a jelszó után

					if (checkPassword(enteredPassword)) {
						String sorszamString = JOptionPane.showInputDialog(null, "Adja meg a törölni kívánt sorszámot:",
								"Törlés", JOptionPane.QUESTION_MESSAGE);
						if (sorszamString != null && !sorszamString.isEmpty()) {
							try {
								int sorszam = Integer.parseInt(sorszamString);
								Row row = getRowBySorszam(sorszam);

								if (row != null) {
									int dialogResult = JOptionPane.showConfirmDialog(null,
											"Biztosan törölni akarja a sort? \n A törlés nem visszavonható!!",
											"Törlés megerősítése", JOptionPane.YES_NO_OPTION);
									if (dialogResult == JOptionPane.YES_OPTION) {
										row.getCell(1).setCellValue("");
										row.getCell(2).setCellValue("");
										row.getCell(3).setCellValue("");
										row.getCell(4).setCellValue("");
										row.getCell(5).setCellValue("");
										row.getCell(6).setCellValue("");

										tableModel.removeRow(row.getRowNum());
										JOptionPane.showMessageDialog(null, "Sikeres törlés!");
									}
								} else {
									JOptionPane.showMessageDialog(null,
											"A megadott sorszám nem létezik az Excel táblázatban!", "Hiba",
											JOptionPane.ERROR_MESSAGE);
								}

								// A változtatások mentése az Excel fájlba
								saveExcelFile();

								// Frissítjük a táblázatot az új adatokkal
								table.setModel(new ExcelTableModel(sheet));
								table.repaint();
							} catch (NumberFormatException ex) {
								JOptionPane.showMessageDialog(null,
										"Érvénytelen sorszám formátum! Adjon meg egy egész számot.", "Hiba",
										JOptionPane.ERROR_MESSAGE);
							}
						}
					} else {
						JOptionPane.showMessageDialog(null, "Helytelen jelszó!", "Hiba", JOptionPane.ERROR_MESSAGE);
					}
				}
			}
		});
		JButton kilepesButton = new JButton("Kilépés");
		kilepesButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				kilepesButtonClicked();
			}
		});
		// Az inputPanel létrehozása és hozzáadása az ablakhoz
		JPanel inputPanel = new JPanel(new GridLayout(5, 6, 5, 5));
		// String today =
		// LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
		Font FontSize = new Font("Arial", Font.BOLD, 18);

		inputPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
		inputPanel.setBackground(new Color(217,217,214));
		inputPanel.add(new JLabel("Belépett(év/hó/nap/óra/perc):"));
		inputPanel.add(beField);
		inputPanel.add(new JLabel("Vendég neve:"));
		inputPanel.add(vendegField);
		inputPanel.add(new JLabel("Kártya száma:"));
		inputPanel.add(kartyaField);
		inputPanel.add(new JLabel("Ügyintéző neve(akihez érkezett):"));
		inputPanel.add(ugyintezoField);
		inputPanel.add(new JLabel("Kilépett(év/hó/nap/óra/perc):"));
		inputPanel.add(kiField);
		inputPanel.add(new JLabel("(Visitor) kártya visszavétele:"));
		inputPanel.add(visitorField);
		inputPanel.add(new JLabel());
		inputPanel.add(new JLabel());
		inputPanel.add(new JLabel());
		inputPanel.add(new JLabel());
		inputPanel.add(hozzaadButton);
		inputPanel.add(modositasButton);
		inputPanel.add(torlesButton);
		inputPanel.add(kilepesButton);

		modositasButton.setForeground(Color.BLUE);
		torlesButton.setForeground(Color.RED);

		hozzaadButton.setFont(FontSize);
		torlesButton.setFont(FontSize);
		kilepesButton.setFont(FontSize);
		modositasButton.setFont(FontSize);

		for (Component component : inputPanel.getComponents()) {
			if (component instanceof JTextField) {
				JTextField textField = (JTextField) component;
				textField.setFont(FontSize);
			}
		}

		for (Component component : inputPanel.getComponents()) {
			if (component instanceof JLabel) {
				JLabel label = (JLabel) component;
				label.setFont(FontSize);
			}
		}
		// Az inputPanel hozzáadása a JFrame-hez
		add(inputPanel, BorderLayout.SOUTH);
		
		pack(); // Méret az aktuális tartalomhoz igazítva
		setLocationRelativeTo(null); // Középre helyezés

		setExtendedState(JFrame.MAXIMIZED_BOTH);

		setVisible(true);
	}

	private Row getRowBySorszam(int sorszam) {
		for (Row row : sheet) {
			if (row.getCell(0) != null && row.getCell(0).getCellType() == CellType.NUMERIC) {
				int cellValue = (int) row.getCell(0).getNumericCellValue();
				if (cellValue == sorszam) {
					return row;
				}
			}
		}
		return null;
	}

	// Az adatokat tároló táblázat modellje
	private class ExcelTableModel extends AbstractTableModel {

		private static final long serialVersionUID = 1L;
		// private Object[][] data; // Adatokat tároló tömb
		private Sheet sheet;

		public ExcelTableModel(Sheet sheet) {
			this.sheet = sheet;
		}

		public void removeRow(int rowIndex) {
			sheet.removeRow(sheet.getRow(rowIndex));
		}

		@Override
		public int getRowCount() {
			return rowCount;
		}

		@Override
		public int getColumnCount() {
			return 7; // 7 oszlopunk van
		}

		@Override
		public Object getValueAt(int rowIndex, int columnIndex) {
			Row row = sheet.getRow(rowIndex);
			if (row == null) {
				return null;
			}
			Cell cell = row.getCell(columnIndex);
			if (cell == null) {
				return null;
			}
			switch (cell.getCellType()) {
			case STRING:
				return cell.getStringCellValue();
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					// For the date column
					return cell.getDateCellValue();
				} else {
					// For other numeric columns
					return (int) cell.getNumericCellValue();
				}
			case BOOLEAN:
				// For boolean columns
				return cell.getBooleanCellValue();
			default:
				return null;
			}
		}

		@Override
		public String getColumnName(int column) {
			String[] oszlopNevek = { "Sorszám:", "Belépett(év/hó/nap/óra/perc):", "Vendég neve:", "Kártya száma:",
					"Ügyintéző neve(akihez érkezett):", "Kilépett(év/hó/nap/óra/perc):",
					"(Visitor) kártya visszavétele:" };

			// Betűméret és stílus beállítása
			int fontSize = 5; // Itt állíthatod be a kívánt betűméretet
			int fontStyle = Font.BOLD; // Itt állíthatod be a kívánt stílust (Font.PLAIN, Font.BOLD, Font.ITALIC, stb.)

			// A megfelelő oszlopnevet visszaadjuk a megadott betűmérettel és stílussal
			return "<html><font face='Arial' size='" + fontSize + "' style='" + fontStyle + "'>" + oszlopNevek[column]
					+ "</font></html>";
		}
	}

	public void saveExcelFile() {
		try {
			String currentDirectory = System.getProperty("user.dir");
			String folderName = "Vendégek_És_Belépők";
			String folderPath = currentDirectory + File.separator + folderName;

			Path folder = Paths.get(folderPath);
			if (!Files.exists(folder)) {
				Files.createDirectory(folder);
			}

			String today = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
			String filePath = folderPath + File.separator + "Vendég_Belépő_" + today + ".xlsx";

			FileOutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
			System.out.println("Excel fájl mentve: " + filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		SwingUtilities.invokeLater(() -> new Vendegek_Es_Belepok());
	}
}