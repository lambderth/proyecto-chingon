package datapull;

import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;

import java.io.FileReader;
import java.util.*;
import java.util.stream.Collectors;

public class ConversationDataProvider {

    public static List<String> loadConversations(String csvPath) {
        LinkedHashMap<String, List<String[]>> groups = new LinkedHashMap<>();
        String currentId = null;

        try (CSVReader reader = new CSVReaderBuilder(new FileReader(csvPath)).build()) {
            List<String[]> rows = reader.readAll();
            if (rows.size() <= 1) return Collections.emptyList(); // no data
            rows.remove(0); // drop header

            for (String[] row : rows) {
                // Expect 3 columns: ID, Visitor, Associate
                String idCell = safeTrim(row, 0);
                if (!idCell.isEmpty()) {
                    currentId = idCell;
                    groups.putIfAbsent(currentId, new ArrayList<>());
                }
                if (currentId == null) continue; // skip stray rows before first ID

                String visitorCell = normalizeNewlines(safeGet(row, 1));
                String associateCell = normalizeNewlines(safeGet(row, 2));

                groups.get(currentId).add(new String[]{visitorCell, associateCell});
            }

        } catch (Exception e) {
            throw new RuntimeException("Failed to read CSV " + csvPath, e);
        }

        // Build ordered conversation string for each group
        return groups.entrySet().stream()
                .map(entry -> {
                    String id = entry.getKey();
                    StringBuilder out = new StringBuilder();
                    out.append("Conversation ").append(id).append(":\n");

                    for (String[] row : entry.getValue()) {
                        String visitor = row[0];
                        String associate = row[1];

                        // For each row keep the order: visitor block then associate block
                        if (!visitor.isBlank()) {
                            appendSpeakerBlock(out, "Visitor", visitor);
                        }
                        if (!associate.isBlank()) {
                            appendSpeakerBlock(out, "Associate", associate);
                        }
                    }

                    return out.toString().trim();
                })
                .collect(Collectors.toList());
    }

    private static void appendSpeakerBlock(StringBuilder sb, String speakerLabel, String text) {
        // Split into lines preserving order
        String[] lines = text.split("\\n", -1);
        if (lines.length == 0) return;
        // Label only the first line; subsequent lines continue without label (as you requested)
        sb.append(speakerLabel).append(": ").append(lines[0]).append("\n");
        for (int i = 1; i < lines.length; i++) {
            sb.append(lines[i]).append("\n");
        }
        sb.append("\n"); // blank line between blocks to improve readability (optional)
    }

    private static String normalizeNewlines(String input) {
        if (input == null) return "";
        // Handle escaped sequences literal "\r\n" or "\n"
        String s = input.replace("\\r\\n", "\n").replace("\\n", "\n");
        // Convert any remaining CRLF or CR into single '\n'
        s = s.replace("\r\n", "\n").replace("\r", "\n");
        return s;
    }

    private static String safeGet(String[] row, int idx) {
        if (row == null || idx >= row.length) return "";
        return row[idx] == null ? "" : row[idx];
    }

    private static String safeTrim(String[] row, int idx) {
        return safeGet(row, idx).trim().replaceAll("^\"|\"$", "");
    }
}
