package com.jira.explorer.service;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.jira.explorer.model.JiraConfig;
import com.jira.explorer.model.JiraIssue;
import okhttp3.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Client for interacting with Jira REST API v2/v3
 * Supports both Jira Server 9.12.24 (API v2) and Jira Cloud (API v3)
 */
public class JiraApiClient {
    private static final Logger logger = LoggerFactory.getLogger(JiraApiClient.class);
    private static final MediaType JSON = MediaType.get("application/json; charset=utf-8");

    private final OkHttpClient httpClient;
    private final JiraConfig config;
    private final Gson gson;

    public JiraApiClient(JiraConfig config) {
        this.config = config;
        this.gson = new Gson();
        this.httpClient = new OkHttpClient.Builder()
                .connectTimeout(30, TimeUnit.SECONDS)
                .readTimeout(60, TimeUnit.SECONDS)
                .writeTimeout(30, TimeUnit.SECONDS)
                .build();
    }

    /**
     * Execute JQL query and return list of issues
     * Uses configured API version (v2 for Server 9.12.24, v3 for Cloud)
     */
    public List<JiraIssue> searchIssues(String jql, int startAt, int maxResults) throws IOException {
        logger.info("Using Jira API version: {}", config.getApiVersion().getDisplayName());
        logger.info("Executing JQL query: {}", jql);

        Request request;

        if (config.getApiVersion() == JiraConfig.ApiVersion.CLOUD_CURRENT) {
            // API v3 (Cloud) uses GET with query parameters on /search/jql endpoint
            String searchEndpoint = config.getApiVersion().getSearchEndpoint();
            HttpUrl.Builder urlBuilder = HttpUrl.parse(config.getJiraUrl() + searchEndpoint).newBuilder();
            urlBuilder.addQueryParameter("jql", jql);
            urlBuilder.addQueryParameter("startAt", String.valueOf(startAt));
            urlBuilder.addQueryParameter("maxResults", String.valueOf(maxResults));
            urlBuilder.addQueryParameter("fields", "*navigable");

            String finalUrl = urlBuilder.build().toString();
            logger.info("Cloud API URL: {}", finalUrl);

            request = new Request.Builder()
                    .url(finalUrl)
                    .get()
                    .addHeader("Authorization", getAuthHeader())
                    .addHeader("Accept", "application/json")
                    .build();
        } else {
            // API v2 (Server) uses POST with JSON body on /search endpoint
            String searchEndpoint = config.getApiVersion().getSearchEndpoint();
            String url = config.getJiraUrl() + searchEndpoint;

            JsonObject requestBody = new JsonObject();
            requestBody.addProperty("jql", jql);
            requestBody.addProperty("startAt", startAt);
            requestBody.addProperty("maxResults", maxResults);
            requestBody.add("fields", new Gson().toJsonTree(new String[]{"*all"}));
            requestBody.addProperty("expand", "names,schema");

            String jsonPayload = gson.toJson(requestBody);
            logger.info("Server API payload: {}", jsonPayload);

            RequestBody body = RequestBody.create(jsonPayload, JSON);

            request = new Request.Builder()
                    .url(url)
                    .post(body)
                    .addHeader("Authorization", getAuthHeader())
                    .addHeader("Accept", "application/json")
                    .addHeader("Content-Type", "application/json")
                    .build();
        }

        try (Response response = httpClient.newCall(request).execute()) {
            if (!response.isSuccessful()) {
                String errorBody = response.body() != null ? response.body().string() : "No error details";
                throw new IOException("Jira API request failed: " + response.code() + " - " + errorBody);
            }

            String responseBody = response.body().string();
            JsonObject jsonResponse = gson.fromJson(responseBody, JsonObject.class);

            List<JiraIssue> issues = new ArrayList<>();
            JsonArray issuesArray = jsonResponse.getAsJsonArray("issues");

            for (int i = 0; i < issuesArray.size(); i++) {
                JsonObject issueJson = issuesArray.get(i).getAsJsonObject();
                issues.add(new JiraIssue(issueJson));
            }

            logger.info("Retrieved {} issues", issues.size());
            return issues;
        }
    }

    /**
     * Test connection to Jira instance using configured API version
     */
    public boolean testConnection() {
        try {
            String apiPath = config.getApiVersion().getApiPath();
            String url = config.getJiraUrl() + apiPath + "/myself";

            Request request = new Request.Builder()
                    .url(url)
                    .get()
                    .addHeader("Authorization", getAuthHeader())
                    .addHeader("Accept", "application/json")
                    .build();

            try (Response response = httpClient.newCall(request).execute()) {
                return response.isSuccessful();
            }
        } catch (Exception e) {
            logger.error("Connection test failed", e);
            return false;
        }
    }

    /**
     * Get field metadata for better field name display using configured API version
     */
    public JsonObject getFieldMetadata() throws IOException {
        String apiPath = config.getApiVersion().getApiPath();
        String url = config.getJiraUrl() + apiPath + "/field";

        Request request = new Request.Builder()
                .url(url)
                .get()
                .addHeader("Authorization", getAuthHeader())
                .addHeader("Accept", "application/json")
                .build();

        try (Response response = httpClient.newCall(request).execute()) {
            if (!response.isSuccessful()) {
                throw new IOException("Failed to fetch field metadata: " + response.code());
            }

            String responseBody = response.body().string();
            JsonArray fieldsArray = gson.fromJson(responseBody, JsonArray.class);

            JsonObject fieldMap = new JsonObject();
            for (int i = 0; i < fieldsArray.size(); i++) {
                JsonObject field = fieldsArray.get(i).getAsJsonObject();
                String id = field.get("id").getAsString();
                String name = field.get("name").getAsString();
                fieldMap.addProperty(id, name);
            }

            return fieldMap;
        }
    }

    private String getAuthHeader() {
        String auth = config.getUsername() + ":" + config.getApiToken();
        String encodedAuth = Base64.getEncoder().encodeToString(auth.getBytes(StandardCharsets.UTF_8));
        return "Basic " + encodedAuth;
    }

    public JiraConfig getConfig() {
        return config;
    }
}
