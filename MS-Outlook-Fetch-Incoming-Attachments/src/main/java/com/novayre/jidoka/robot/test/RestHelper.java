package com.novayre.jidoka.robot.test;

import java.io.IOException;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
public class RestHelper {
  /**
   * Converts a Java object to JSON
   *
   * @param obj Any object to be serialized to JSON
   * @return JSON serialization of obj
   */
  static String toJson(Object obj) {
    ObjectMapper objectMapper = new ObjectMapper();
    String json = "";
    try {
      json = objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(obj);
    } catch (JsonProcessingException e) {
      e.printStackTrace();
    }
    return json;
  }


  /**
   * Deserializes a JSON string into the type provided as typeRef.
   *
   * Example:
   *
   * Map<String, Object> map = RestHelper.fromJson(someJson, new TypeReference<Map<String, Object>>(){});
   *
   * @param json JSON serialization string
   */
  static <T> Object fromJson(String json) {
    ObjectMapper objectMapper = new ObjectMapper();
    T value = null;
    try {
      value = objectMapper.readValue(json, new TypeReference<T>(){});
    } catch (IOException e) {
      e.printStackTrace();
    }
    return value;
  }
}
