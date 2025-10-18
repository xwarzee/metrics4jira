package com.jira.explorer.model;

import com.google.gson.JsonElement;
import com.google.gson.JsonObject;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Model class representing a Jira issue with all its fields
 */
public class JiraIssue {
    private String key;
    private String id;
    private String self;
    private JsonObject fields;
    private Map<String, Object> flattenedFields;

    public JiraIssue(JsonObject issueJson) {
        this.key = issueJson.get("key").getAsString();
        this.id = issueJson.get("id").getAsString();
        this.self = issueJson.get("self").getAsString();
        this.fields = issueJson.getAsJsonObject("fields");
        this.flattenedFields = new LinkedHashMap<>();
        flattenFields();
    }

    private void flattenFields() {
        if (fields == null) return;

        for (Map.Entry<String, JsonElement> entry : fields.entrySet()) {
            String fieldName = entry.getKey();
            JsonElement value = entry.getValue();

            if (value.isJsonNull()) {
                flattenedFields.put(fieldName, null);
            } else if (value.isJsonPrimitive()) {
                flattenedFields.put(fieldName, value.getAsString());
            } else if (value.isJsonObject()) {
                JsonObject obj = value.getAsJsonObject();
                // Try to get common display fields
                if (obj.has("displayName")) {
                    flattenedFields.put(fieldName, obj.get("displayName").getAsString());
                } else if (obj.has("name")) {
                    flattenedFields.put(fieldName, obj.get("name").getAsString());
                } else if (obj.has("value")) {
                    flattenedFields.put(fieldName, obj.get("value").getAsString());
                } else {
                    flattenedFields.put(fieldName, obj.toString());
                }
            } else if (value.isJsonArray()) {
                flattenedFields.put(fieldName, value.toString());
            }
        }
    }

    public String getKey() {
        return key;
    }

    public String getId() {
        return id;
    }

    public String getSelf() {
        return self;
    }

    public JsonObject getFields() {
        return fields;
    }

    public Map<String, Object> getFlattenedFields() {
        return flattenedFields;
    }

    public String getFieldValue(String fieldName) {
        Object value = flattenedFields.get(fieldName);
        return value != null ? value.toString() : "";
    }

    public String getSummary() {
        return getFieldValue("summary");
    }

    public String getStatus() {
        if (fields.has("status") && !fields.get("status").isJsonNull()) {
            JsonObject status = fields.getAsJsonObject("status");
            if (status.has("name")) {
                return status.get("name").getAsString();
            }
        }
        return "";
    }

    public String getIssueType() {
        if (fields.has("issuetype") && !fields.get("issuetype").isJsonNull()) {
            JsonObject issueType = fields.getAsJsonObject("issuetype");
            if (issueType.has("name")) {
                return issueType.get("name").getAsString();
            }
        }
        return "";
    }

    public String getAssignee() {
        if (fields.has("assignee") && !fields.get("assignee").isJsonNull()) {
            JsonObject assignee = fields.getAsJsonObject("assignee");
            if (assignee.has("displayName")) {
                return assignee.get("displayName").getAsString();
            }
        }
        return "Unassigned";
    }

    @Override
    public String toString() {
        return key + " - " + getSummary();
    }
}
