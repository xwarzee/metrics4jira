package com.jira.explorer.ui;

import com.jira.explorer.model.JiraConfig;
import javafx.geometry.Insets;
import javafx.scene.control.*;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;

import java.util.Optional;

/**
 * Dialog for configuring Jira connection settings
 */
public class ConfigDialog extends Dialog<JiraConfig> {

    private final TextField jiraUrlField;
    private final TextField usernameField;
    private final PasswordField apiTokenField;
    private final TextField maxResultsField;
    private final ComboBox<JiraConfig.ApiVersion> apiVersionComboBox;

    public ConfigDialog(JiraConfig currentConfig) {
        setTitle("Configure Jira Connection");
        setHeaderText("Enter your Jira instance details");

        // Set the button types
        ButtonType saveButtonType = new ButtonType("Save", ButtonBar.ButtonData.OK_DONE);
        getDialogPane().getButtonTypes().addAll(saveButtonType, ButtonType.CANCEL);

        // Create the form
        GridPane grid = new GridPane();
        grid.setHgap(10);
        grid.setVgap(10);
        grid.setPadding(new Insets(20, 150, 10, 10));

        jiraUrlField = new TextField();
        jiraUrlField.setPromptText("https://your-domain.atlassian.net");
        jiraUrlField.setPrefWidth(400);
        if (currentConfig != null && currentConfig.getJiraUrl() != null) {
            jiraUrlField.setText(currentConfig.getJiraUrl());
        }

        usernameField = new TextField();
        usernameField.setPromptText("your-email@example.com");
        if (currentConfig != null && currentConfig.getUsername() != null) {
            usernameField.setText(currentConfig.getUsername());
        }

        apiTokenField = new PasswordField();
        apiTokenField.setPromptText("Your Jira API token");
        if (currentConfig != null && currentConfig.getApiToken() != null) {
            apiTokenField.setText(currentConfig.getApiToken());
        }

        maxResultsField = new TextField();
        maxResultsField.setPromptText("50");
        maxResultsField.setText(currentConfig != null ? String.valueOf(currentConfig.getMaxResults()) : "50");
        maxResultsField.setPrefWidth(100);

        // API Version ComboBox
        apiVersionComboBox = new ComboBox<>();
        apiVersionComboBox.getItems().addAll(JiraConfig.ApiVersion.values());
        apiVersionComboBox.setValue(currentConfig != null ? currentConfig.getApiVersion() : JiraConfig.ApiVersion.CLOUD_CURRENT);
        apiVersionComboBox.setCellFactory(param -> new ListCell<JiraConfig.ApiVersion>() {
            @Override
            protected void updateItem(JiraConfig.ApiVersion item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) {
                    setText(null);
                } else {
                    setText(item.getDisplayName());
                }
            }
        });
        apiVersionComboBox.setButtonCell(new ListCell<JiraConfig.ApiVersion>() {
            @Override
            protected void updateItem(JiraConfig.ApiVersion item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) {
                    setText(null);
                } else {
                    setText(item.getDisplayName());
                }
            }
        });
        apiVersionComboBox.setPrefWidth(400);

        // Add labels and fields to grid
        grid.add(new Label("Jira URL:"), 0, 0);
        grid.add(jiraUrlField, 1, 0);

        grid.add(new Label("API Version:"), 0, 1);
        grid.add(apiVersionComboBox, 1, 1);

        grid.add(new Label("Username (Email):"), 0, 2);
        grid.add(usernameField, 1, 2);

        grid.add(new Label("API Token:"), 0, 3);
        grid.add(apiTokenField, 1, 3);

        grid.add(new Label("Max Results:"), 0, 4);
        grid.add(maxResultsField, 1, 4);

        // Add info label
        Label infoLabel = new Label(
            "API Version:\n" +
            "- Jira Server 9.12.24: Use API v2 for on-premise Jira Server installations\n" +
            "- Jira Cloud (Current): Use API v3 for Atlassian Cloud instances\n\n" +
            "To generate an API token:\n" +
            "1. Go to https://id.atlassian.com/manage-profile/security/api-tokens\n" +
            "2. Click 'Create API token'\n" +
            "3. Copy the token and paste it here"
        );
        infoLabel.setStyle("-fx-font-size: 10; -fx-text-fill: #666;");
        infoLabel.setWrapText(true);
        infoLabel.setMaxWidth(400);
        grid.add(infoLabel, 1, 5);

        getDialogPane().setContent(grid);

        // Enable/Disable save button depending on whether fields are filled
        Button saveButton = (Button) getDialogPane().lookupButton(saveButtonType);
        saveButton.addEventFilter(javafx.event.ActionEvent.ACTION, event -> {
            if (!validateInput()) {
                event.consume();
                showValidationError();
            }
        });

        // Convert the result when save button is clicked
        setResultConverter(dialogButton -> {
            if (dialogButton == saveButtonType) {
                JiraConfig config = new JiraConfig();
                config.setJiraUrl(jiraUrlField.getText().trim());
                config.setUsername(usernameField.getText().trim());
                config.setApiToken(apiTokenField.getText().trim());
                config.setApiVersion(apiVersionComboBox.getValue());
                try {
                    config.setMaxResults(Integer.parseInt(maxResultsField.getText().trim()));
                } catch (NumberFormatException e) {
                    config.setMaxResults(50);
                }
                return config;
            }
            return null;
        });
    }

    private boolean validateInput() {
        String url = jiraUrlField.getText().trim();
        String username = usernameField.getText().trim();
        String token = apiTokenField.getText().trim();

        if (url.isEmpty() || username.isEmpty() || token.isEmpty()) {
            return false;
        }

        // Validate URL format
        if (!url.startsWith("http://") && !url.startsWith("https://")) {
            return false;
        }

        // Validate max results is a number
        try {
            int maxResults = Integer.parseInt(maxResultsField.getText().trim());
            if (maxResults < 1 || maxResults > 1000) {
                return false;
            }
        } catch (NumberFormatException e) {
            return false;
        }

        return true;
    }

    private void showValidationError() {
        Alert alert = new Alert(Alert.AlertType.ERROR);
        alert.setTitle("Validation Error");
        alert.setHeaderText("Invalid Configuration");
        alert.setContentText(
            "Please ensure:\n" +
            "- Jira URL starts with http:// or https://\n" +
            "- Username and API Token are not empty\n" +
            "- Max Results is a number between 1 and 1000"
        );
        alert.showAndWait();
    }
}
