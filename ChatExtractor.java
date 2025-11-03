package datapull;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.opencsv.CSVWriter;

import java.io.*;
import java.util.*;

public class ChatExtractor {

    public static void main(String[] args) {
        String inputFile = "input.xlsx";   // Your input Excel file
        String outputFile = "output.csv";  // Output CSV file

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(fis);
             FileWriter fw = new FileWriter(outputFile);
             CSVWriter writer = new CSVWriter(fw,
                     CSVWriter.DEFAULT_SEPARATOR,
                     CSVWriter.DEFAULT_QUOTE_CHARACTER,   // Ensures double quotes
                     CSVWriter.DEFAULT_ESCAPE_CHARACTER,
                     CSVWriter.DEFAULT_LINE_END)) {

            Sheet sheet = workbook.getSheetAt(0);
            boolean afterRobot = false;
            boolean skipHeader = true;

            List<String[]> rows = new ArrayList<>();
            StringBuilder visitorMsg = new StringBuilder();
            StringBuilder sender2Msg = new StringBuilder();
            String currentSender = "";
            String sender2Name = "Associate"; // Default until detected

            for (Row row : sheet) {
                if (skipHeader) {
                    skipHeader = false;
                    continue; // skip Excel header row
                }

                Cell cell = row.getCell(7); // Column H (index 7)
                if (cell == null) continue;

                String text = cell.toString().trim();
                if (text.isEmpty()) continue;

                // Wait until the Robot message is found
                if (!afterRobot) {
                    if (text.contains("Robot: Provide a quick summary of your request")) {
                        afterRobot = true;
                    }
                    continue;
                }

                // Parse sender and message
                int colonIndex = text.indexOf(":");
                if (colonIndex == -1) continue;

                String senderPart = text.substring(0, colonIndex).trim();
                String message = text.substring(colonIndex + 1).trim();

                // Remove timestamp e.g. "( 1m 24s )"
                if (senderPart.contains(")")) {
                    senderPart = senderPart.substring(senderPart.indexOf(")") + 1).trim();
                }

                // Skip Robot lines
                if (senderPart.equalsIgnoreCase("Robot")) continue;

                boolean isVisitor = senderPart.equalsIgnoreCase("Visitor");

                // Capture the first non-Visitor name as the associate name
                if (!isVisitor && sender2Name.equals("Associate")) {
                    sender2Name = senderPart;
                }

                // If sender changes, flush previous messages
                if (!currentSender.isEmpty() && !currentSender.equals(senderPart)) {
                    rows.add(new String[]{
                            visitorMsg.toString().trim(),
                            sender2Msg.toString().trim()
                    });
                    visitorMsg.setLength(0);
                    sender2Msg.setLength(0);
                }

                currentSender = senderPart;

                // Add message text (no timestamp, no name)
                if (isVisitor) {
                    if (visitorMsg.length() > 0) visitorMsg.append("\n");
                    visitorMsg.append(message);
                } else {
                    if (sender2Msg.length() > 0) sender2Msg.append("\n");
                    sender2Msg.append(message);
                }
            }

            // Add last accumulated messages
            if (visitorMsg.length() > 0 || sender2Msg.length() > 0) {
                rows.add(new String[]{
                        visitorMsg.toString().trim(),
                        sender2Msg.toString().trim()
                });
            }

            // Write CSV header and rows
            writer.writeNext(new String[]{"Visitor", sender2Name});
            for (String[] row : rows) {
                writer.writeNext(new String[]{
                        row[0].replaceAll("\r", ""),
                        row[1].replaceAll("\r", "")
                });
            }

            System.out.println("✅ CSV created successfully: " + outputFile);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
