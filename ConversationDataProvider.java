package datapull;

import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import java.io.FileReader;
import java.util.*;
import java.util.regex.*;
import java.util.stream.Collectors;

public class ConversationDataProvider {

    public static class Section {
        public String sectionId;
        public List<String> visitorMessages;
        public List<String> associateMessages;

        public Section(String sectionId, List<String> visitorMessages, List<String> associateMessages) {
            this.sectionId = sectionId;
            this.visitorMessages = visitorMessages;
            this.associateMessages = associateMessages;
        }

        @Override
        public String toString() {
            return String.format("{ \"%s\", %s, %s }", sectionId, visitorMessages, associateMessages);
        }
    }

    /**
     * Returns a list of conversations. Each conversation is a list of Sections.
     */
    public static List<List<Section>> loadConversations(String csvPath) {
        LinkedHashMap<String, List<Section>> conversationsMap = new LinkedHashMap<>();

        try (CSVReader reader = new CSVReaderBuilder(new FileReader(csvPath)).build()) {
            List<String[]> rows = reader.readAll();
            if (rows.size() <= 1) return Collections.emptyList();
            rows.remove(0); // remove header

            for (String[] row : rows) {
                String rawId = safeTrim(row, 0);
                if (rawId.isEmpty()) continue;

                String conversationId = extractConversationId(rawId); // e.g. "C1" from "C1 S2"
                String sectionId = rawId;

                String visitorRaw = safeGet(row, 1);
                String associateRaw = safeGet(row, 2);

                List<String> visitorMsgs = splitMessages(visitorRaw);
                List<String> associateMsgs = splitMessages(associateRaw);

                Section section = new Section(sectionId, visitorMsgs, associateMsgs);
                conversationsMap.computeIfAbsent(conversationId, k -> new ArrayList<>()).add(section);
            }

        } catch (Exception e) {
            throw new RuntimeException("Failed to parse CSV: " + e.getMessage(), e);
        }

        // Convert to ordered list of conversations
        return new ArrayList<>(conversationsMap.values());
    }

    // --- Utility methods ---

    private static String extractConversationId(String rawId) {
        // Capture prefix like "C1" from "C1 S1"
        Matcher m = Pattern.compile("^(C\\d+)").matcher(rawId.trim());
        return m.find() ? m.group(1) : rawId;
    }

    private static List<String> splitMessages(String text) {
        if (text == null || text.isBlank()) return Collections.singletonList("");
        // Normalize any escaped \r\n or \n and then split
        String normalized = text.replace("\\r\\n", "\n").replace("\\n", "\n").replace("\r\n", "\n").replace("\r", "\n");
        return Arrays.stream(normalized.split("\\n"))
                .map(String::trim)
                .filter(s -> !s.isEmpty())
                .collect(Collectors.toList());
    }

    private static String safeGet(String[] row, int idx) {
        if (row == null || idx >= row.length) return "";
        return row[idx] == null ? "" : row[idx];
    }

    private static String safeTrim(String[] row, int idx) {
        return safeGet(row, idx).trim().replaceAll("^\"|\"$", "");
    }
}
