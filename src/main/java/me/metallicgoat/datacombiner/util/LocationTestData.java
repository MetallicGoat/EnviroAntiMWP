package me.metallicgoat.datacombiner.util;

import java.util.Date;
import java.util.HashMap;

public class LocationTestData {

  // Analyte, Value
  public HashMap<String, String> data = new HashMap<>();

  // Date the location was collected
  public Date dataDate;


  public void addData(String param, String value) {
    data.put(ParamTranslator.standardizeParam(param), value);
  }

  public String getDataByParam(String param) {
    param = ParamTranslator.standardizeParam(param);

    final String result = data.get(param);

    if (result == null) {
      // Try to get the standardized name from Param enum
      final String standardizedParam = ParamTranslator.getStandardName(param);
      return data.get(standardizedParam);
    }

    return result;
  }
}
