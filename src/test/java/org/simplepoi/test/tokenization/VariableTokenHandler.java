package org.simplepoi.test.tokenization;

import java.util.Properties;

public class VariableTokenHandler implements TokenHandler {

    private final String openToken; // #{  ${
    private final String closeToken; // }
    private final Properties variables;

    public VariableTokenHandler(Properties variables, String openToken, String closeToken) {
        this.variables = variables;
        this.openToken = openToken;
        this.closeToken = closeToken;
    }

    private String getPropertyValue(String key, String defaultValue) {
        return (variables == null) ? defaultValue : variables.getProperty(key, defaultValue);
    }

    @Override
    public String handleToken(String content) {
        if (variables != null && variables.containsKey(content)) {
            return variables.getProperty(content);
        }
        return this.openToken+ content + this.closeToken;
    }
}
