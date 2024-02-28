package org.simplepoi.excel.export;

import java.util.Properties;

public interface TokenHandler {
  String handleToken(String content);

   void setSecondProp(Properties prop);
}


