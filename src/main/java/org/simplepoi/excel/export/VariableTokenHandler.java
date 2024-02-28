package org.simplepoi.excel.export;

import java.util.Properties;

public class VariableTokenHandler implements TokenHandler {

    // refresh property value todo
    private final String openToken; // #{  ${
    private final String closeToken; // }
    private final Properties variables;
    private   Properties variables2;

    public VariableTokenHandler(Properties variables, String openToken, String closeToken) {
        this.variables = variables;
        this.openToken = openToken;
        this.closeToken = closeToken;
    }

    private String getPropertyValue(String key, String defaultValue) {
        return (variables == null) ? defaultValue : variables.getProperty(key, defaultValue);
    }

    public void setSecondProp(Properties prop) {
     this.variables2 = prop;
    }

    @Override
    public String handleToken(String content) {
        if (variables2 != null && variables2.containsKey(content)) {
            return variables2.getProperty(content);
        }
        if (variables != null && variables.containsKey(content)) {
            return variables.getProperty(content);
        }
        return this.openToken+ content + this.closeToken;
    }
}
