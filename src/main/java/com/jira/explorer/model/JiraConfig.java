package com.jira.explorer.model;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

/**
 * Configuration class for Jira connection settings
 */
public class JiraConfig {

    /**
     * Enum for Jira API versions
     */
    public enum ApiVersion {
        SERVER_9_12_24("Jira Server 9.12.24", "2", "/rest/api/2"),
        CLOUD_CURRENT("Jira Cloud (Current)", "3", "/rest/api/3");

        private final String displayName;
        private final String apiVersion;
        private final String apiPath;

        ApiVersion(String displayName, String apiVersion, String apiPath) {
            this.displayName = displayName;
            this.apiVersion = apiVersion;
            this.apiPath = apiPath;
        }

        public String getDisplayName() {
            return displayName;
        }

        public String getApiVersion() {
            return apiVersion;
        }

        public String getApiPath() {
            return apiPath;
        }

        public String getSearchEndpoint() {
            // Cloud v3 uses /search/jql, Server v2 uses /search
            if (this == CLOUD_CURRENT) {
                return apiPath + "/search/jql";
            }
            return apiPath + "/search";
        }
    }

    private String jiraUrl;
    private String username;
    private String apiToken;
    private int maxResults;
    private ApiVersion apiVersion;

    public JiraConfig() {
        loadFromProperties();
    }

    public JiraConfig(String jiraUrl, String username, String apiToken) {
        this.jiraUrl = jiraUrl;
        this.username = username;
        this.apiToken = apiToken;
        this.maxResults = 50; // Default
        this.apiVersion = ApiVersion.CLOUD_CURRENT; // Default
    }

    private void loadFromProperties() {
        Properties props = new Properties();
        try (InputStream input = getClass().getClassLoader().getResourceAsStream("jira.properties")) {
            if (input != null) {
                props.load(input);
                this.jiraUrl = props.getProperty("jira.url", "");
                this.username = props.getProperty("jira.username", "");
                this.apiToken = props.getProperty("jira.apitoken", "");
                this.maxResults = Integer.parseInt(props.getProperty("jira.maxresults", "50"));

                // Load API version
                String versionStr = props.getProperty("jira.apiversion", "CLOUD_CURRENT");
                try {
                    this.apiVersion = ApiVersion.valueOf(versionStr);
                } catch (IllegalArgumentException e) {
                    this.apiVersion = ApiVersion.CLOUD_CURRENT;
                }
            } else {
                this.maxResults = 50;
                this.apiVersion = ApiVersion.CLOUD_CURRENT;
            }
        } catch (IOException e) {
            // Properties file not found or error reading, use defaults
            this.maxResults = 50;
            this.apiVersion = ApiVersion.CLOUD_CURRENT;
        }
    }

    public String getJiraUrl() {
        return jiraUrl;
    }

    public void setJiraUrl(String jiraUrl) {
        this.jiraUrl = jiraUrl;
    }

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public String getApiToken() {
        return apiToken;
    }

    public void setApiToken(String apiToken) {
        this.apiToken = apiToken;
    }

    public int getMaxResults() {
        return maxResults;
    }

    public void setMaxResults(int maxResults) {
        this.maxResults = maxResults;
    }

    public ApiVersion getApiVersion() {
        return apiVersion != null ? apiVersion : ApiVersion.CLOUD_CURRENT;
    }

    public void setApiVersion(ApiVersion apiVersion) {
        this.apiVersion = apiVersion;
    }

    public boolean isValid() {
        return jiraUrl != null && !jiraUrl.isEmpty() &&
               username != null && !username.isEmpty() &&
               apiToken != null && !apiToken.isEmpty();
    }
}
