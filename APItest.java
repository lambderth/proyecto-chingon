import net.serenitybdd.rest.SerenityRest;
import io.restassured.response.Response;
import org.junit.Test;

public class EinsteinRepliesAPITest {

    private static final String INSTANCE_URL = "https://yourInstance.salesforce.com"; // e.g. https://mydomain.my.salesforce.com
    private static final String API_VERSION = "v64.0"; // or whatever version your org supports
    private static final String PROMPT_TEMPLATE = "einstein_gpt_serviceRepliesGroundedLiveChatTranscript";
    private static final String ACCESS_TOKEN = "YOUR_SF_ACCESS_TOKEN";

    @Test
    public void testEinsteinServiceReplies() {
        // Build endpoint URL
        String endpoint = String.format(
            "%s/services/data/%s/einstein/prompt-templates/%s/generations",
            INSTANCE_URL, API_VERSION, PROMPT_TEMPLATE
        );

        // Example request body (adapt this to your prompt template inputs)
        String requestBody = """
        {
          "input": {
            "chatTranscript": "Customer: Hi, I need help with my recent order.\nAgent: Sure, could you please share your order number?",
            "caseId": "500xx00000123ABC",
            "language": "en_US"
          },
          "maxTokens": 200,
          "temperature": 0.7
        }
        """;

        // Send POST request
        Response response = SerenityRest.given()
            .baseUri(INSTANCE_URL)
            .header("Authorization", "Bearer " + ACCESS_TOKEN)
            .header("Content-Type", "application/json")
            .body(requestBody)
            .post(String.format("/services/data/%s/einstein/prompt-templates/%s/generations",
                    API_VERSION, PROMPT_TEMPLATE));

        // Print and validate
        response.then().log().all();
        response.then().statusCode(200);

        String generatedReply = response.jsonPath().getString("output[0].content");
        System.out.println("Generated Reply:\n" + generatedReply);
    }
}
