package me.metallicgoat.datacombiner;

import java.util.Date;
import java.util.HashMap;

public class LocationTestData {

  // Analyte, Value
  public HashMap<String, String> data = new HashMap<>();

  // Date the location was collected
  public Date dataDate;


  public void addData(String param, String value) {
    data.put(Param.standardizeParam(param), value);
  }

  public String getDataByParam(String param) {
    param = Param.standardizeParam(param);

    final String result = data.get(param);

    if (result == null) {
      // Try to get the standardized name from Param enum
      final String standardizedParam = Param.getStandardName(param);
      return data.get(standardizedParam);
    }

    return result;
  }
}
