package csvformat;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.opencsv.CSVWriter;

import java.io.*;
import java.util.*;

public class ChatConversationExtractor {

    static class ConversationRow {
        String id;
        String visitor = "";
        String associateOrBot = "";

        ConversationRow(String id) {
            this.id = id;
        }
    }

    public static void main(String[] args) {
        String inputPath = "Libro1.xlsx";   // path to your Excel
        String outputPath = "output.csv";  // output CSV path

        try (FileInputStream fis = new FileInputStream(inputPath);
             Workbook workbook = new XSSFWorkbook(fis);
             FileWriter fw = new FileWriter(outputPath);
             CSVWriter csvWriter = new CSVWriter(fw)) {

            Sheet sheet = workbook.getSheetAt(0);

            // CSV header
            csvWriter.writeNext(new String[]{"ID", "Visitor", "Associate/Chatbot"});

            int conversationIndex = 1;

            for (Row row : sheet) {
                // Skip header row (first row)
                if (row.getRowNum() == 0) continue;
                
                Cell cell = row.getCell(7); // column H (0-based)
                if (cell == null) continue;

                String text = cell.toString().trim();
                if (text.isEmpty()) continue;

                String[] lines = text.split("\\r?\\n");
                List<ConversationRow> rows = pairMessages(lines, conversationIndex);

                for (ConversationRow cr : rows) {
                    csvWriter.writeNext(new String[]{cr.id, cr.visitor, cr.associateOrBot});
                }

                conversationIndex++;
            }

            System.out.println("✅ CSV created successfully: " + outputPath);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static List<ConversationRow> pairMessages(String[] lines, int conversationIndex) {
        List<ConversationRow> output = new ArrayList<>();

        List<String> visitorBuffer = new ArrayList<>();
        List<String> botBuffer = new ArrayList<>();

        int step = 1;

        for (String line : lines) {
            line = line.trim();
            if (line.isEmpty()) continue;

            // Check if line has a timestamp - must start with "( " and contain ") "
            if (!line.startsWith("( ") || !line.contains(") ")) {
                // Skip messages without timestamp
                continue;
            }

            // Remove timestamp up to first ") "
            int closeParen = line.indexOf(") ");
            if (closeParen != -1)
                line = line.substring(closeParen + 2).trim();

            String sender = getSender(line);

            if (sender.equals("Visitor")) {
                // If we have a visitor message already and bot messages, flush the row
                // (new visitor message means we're starting a new conversation turn)
                if (!visitorBuffer.isEmpty() && !botBuffer.isEmpty()) {
                    output.add(makeRow(conversationIndex, step++, visitorBuffer, botBuffer));
                    visitorBuffer.clear();
                    botBuffer.clear();
                }
                // If we only have bot messages (no visitor), flush them first
                else if (!botBuffer.isEmpty() && visitorBuffer.isEmpty()) {
                    output.add(makeRow(conversationIndex, step++, visitorBuffer, botBuffer));
                    botBuffer.clear();
                }
                // Add visitor message (consecutive visitors will be grouped)
                visitorBuffer.add(line);
            } else {
                // Non-visitor message (Bot or Associate)
                // Add to bot buffer (will be paired with current visitor or stand alone)
                botBuffer.add(line);
            }
        }

        // Flush any remaining buffers
        if (!visitorBuffer.isEmpty() || !botBuffer.isEmpty()) {
            output.add(makeRow(conversationIndex, step, visitorBuffer, botBuffer));
        }

        return output;
    }

    private static ConversationRow makeRow(int conversationIndex, int step,
                                           List<String> visitorMsgs, List<String> botMsgs) {
        ConversationRow row = new ConversationRow("C" + conversationIndex + " S" + step);
        if (!visitorMsgs.isEmpty())
            row.visitor = String.join("\r\n", visitorMsgs);
        if (!botMsgs.isEmpty())
            row.associateOrBot = String.join("\r\n", botMsgs);
        return row;
    }

    private static String getSender(String msg) {
        int colonIndex = msg.indexOf(":");
        if (colonIndex != -1) {
            String sender = msg.substring(0, colonIndex).trim();
            // Normalize known types
            if (sender.equalsIgnoreCase("Ask Ed Digital Assistant")) return "Bot";
            if (sender.equalsIgnoreCase("Visitor")) return "Visitor";
            return "Associate"; // any other name = real human associate
        }
        return "Unknown";
    }
}
