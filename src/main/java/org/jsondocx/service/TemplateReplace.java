package org.jsondocx.service;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONException;
import com.alibaba.fastjson.JSONObject;
import org.jsondocx.exception.CommonException;
import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.core.VariablePattern;
import pl.jsolve.templ4docx.variable.*;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 模板参数替换
 *
 * @author wwyx778 2019年5月10日 10:44:16
 */
public class TemplateReplace {

    private static final String PREFIX = "#{";
    private static final String SUFFIX = "}";

    /**
     * 生成表格变量
     *
     * @param tableKey       参数
     * @param tableDataArray 需要替换的行数据
     * @throws JSONException JSONException
     */
    private TableVariable generateTableVariables(String tableKey, JSONArray tableDataArray)
            throws JSONException {
        TableVariable tableVariable = new TableVariable();
        //先从第一行数据取到所有的key即所有的列
        List<String> keys = new ArrayList<>(tableDataArray.getJSONObject(0).keySet());
        //校验key数量（表格列数量是否一致）
        int checkKeyCount = keys.size();
        for (String key : keys) {
            List<Variable> columnVariables = new ArrayList<>();
            //一个Key对应一个columnVariables
            if (!tableDataArray.isEmpty()) {
                for (Object rowDataObj : tableDataArray) {
                    //rowDataObj是一行数据
                    if (rowDataObj instanceof JSONObject) {
                        JSONObject rowDataJson = (JSONObject) rowDataObj;
                        //每次循环校验rowDataJson的键数量即表格列数量是否等于从第一行取出的列数量
                        if (checkKeyCount == rowDataJson.keySet().size()) {
                            columnVariables.add(new TextVariable(handleKey(tableKey + "." + key), rowDataJson.getString(key)));
                        } else {
                            throw new CommonException("要替换的表格数据列数不一致！");
                        }
                    }
                }
            }
            tableVariable.addVariable(columnVariables);
        }
        return tableVariable;
    }

    /**
     * 生成图片变量
     *
     * @param imageJson 图片变量替换数据
     * @throws JSONException JSONException
     */
    private ImageVariable generateImageVariables(JSONObject imageJson) throws JSONException {
        Object width = imageJson.get("width");
        Object height = imageJson.get("height");
        if (width instanceof Integer && height instanceof Integer) {
            try {
                return new ImageVariable(handleKey(imageJson.getString("imageKey")), imageJson.getString("imagePath"), (Integer) width, (Integer) height);
            } catch (Exception e) {
                throw new CommonException("图片读取失败！");
            }
        } else {
            throw new CommonException("图片宽度或高度无法识别！");
        }
    }

    /**
     * 处理前缀后缀
     *
     * @param key 替换关键字
     * @return String 有前缀后缀的替换关键字
     */
    private String handleKey(String key) {
        return PREFIX + key + SUFFIX;
    }

    /**
     * 模板参数替换
     *
     * @param inputStream      模板文件输入流
     * @param outputFileName   输出文件名
     * @param templateVariable 模板替换参数
     * @return String 新文件的UUID
     */
    public byte[] templateReplace(InputStream inputStream, String outputFileName, JSONObject templateVariable) {
        Docx templateDocx = new Docx(inputStream);
        Variables variables = new Variables();
        templateDocx.setVariablePattern(new VariablePattern(PREFIX, SUFFIX));

        Object tableJsonArray = templateVariable.get("tableVariable");
        Object imageJsonArray = templateVariable.get("imageVariable");
        Object stringJson = templateVariable.get("variable");

        //表格参数替换
        if (tableJsonArray instanceof JSONArray) {
            //待替换参数的数组
            JSONArray jsonArray = (JSONArray) tableJsonArray;
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject tableJson = jsonArray.getJSONObject(i);
                variables.addTableVariable(generateTableVariables(tableJson.getString("tableKey"), tableJson.getJSONArray("tableData")));
            }
        }

        //图片参数替换
        if (imageJsonArray instanceof JSONArray) {
            //待替换参数的数组
            JSONArray jsonArray = (JSONArray) imageJsonArray;
            for (Object object : jsonArray) {
                //一个imageJson对应一个图片变量
                if (object instanceof JSONObject) {
                    JSONObject imageJson = (JSONObject) object;
                    variables.addImageVariable(generateImageVariables(imageJson));
                }
            }
        }

        //单字段参数替换
        if (stringJson instanceof JSONObject) {
            //待替换参数
            JSONObject jsonObject = (JSONObject) stringJson;
            List<String> keyList = new ArrayList<>(jsonObject.keySet());
            for (String key : keyList) {
                variables.addTextVariable(new TextVariable(handleKey(key), jsonObject.getString(key)));
                System.out.print("name:" + handleKey(key) + "------>");
                System.out.println("value:" + jsonObject.getString(key));
            }
        }

        // fill template
        templateDocx.fillTemplate(variables);
        ByteArrayOutputStream bao = new ByteArrayOutputStream();
        templateDocx.save(bao);
        return bao.toByteArray();
    }

    /**
     * 生成模板变量结构
     *
     * @param stringVariable 单字段变量
     * @param imageVariable  图片变量
     * @param tableVariable  表格变量
     * @return JSONObject 模板变量结构
     */
    public JSONObject getTemplateVariable(JSONObject stringVariable, JSONArray imageVariable, JSONArray tableVariable) {
        JSONObject templateVariables = new JSONObject();
        if (stringVariable != null) {
            templateVariables.put("variable", stringVariable);
        }
        if (imageVariable != null) {
            templateVariables.put("imageVariable", imageVariable);
        }
        if (tableVariable != null) {
            templateVariables.put("tableVariable", tableVariable);
        }
        return templateVariables;
    }

    /**
     * 生成表格变量结构
     *
     * @param map Map<String, ArrayList<?>>
     *            String :表格标识（替换关键字）
     *            ArrayList<?> :表格数据
     * @return JSONObject 模板变量结构
     */
    public JSONArray getTableVariable(Map<String, List<?>> map) {
        JSONArray tableVariable = new JSONArray();
        List<String> keyList = new ArrayList<>(map.keySet());
        for (String key : keyList) {
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("tableKey", key);
            jsonObject.put("tableData", (map.get(key)));
            tableVariable.add(jsonObject);
        }
        return tableVariable;
    }

    /**
     * 生成图片变量结构
     *
     * @param imageKey  图片标识（替换关键字）
     * @param imagePath 图片绝对路径
     * @param width     图片宽度
     * @param height    图片高度
     * @return JSONObject 模板变量结构
     */
    public JSONObject getImageVariable(String imageKey, String imagePath, Integer width, Integer height) {
        JSONObject jsonObject = new JSONObject();
        jsonObject.put("imageKey", imageKey);
        jsonObject.put("imagePath", imagePath);
        jsonObject.put("width", width);
        jsonObject.put("height", height);
        return jsonObject;
    }

    /**
     * 生成单字段变量结构
     *
     * @param map Map<String, String>
     *            String :替换标识（替换关键字）
     *            String :替换数据
     * @return JSONObject 模板变量结构
     */
    public JSONObject getStringVariable(Map<String, String> map) {
        JSONObject jsonObject = new JSONObject();
        List<String> keyList = new ArrayList<>(map.keySet());
        for (String key : keyList) {
            jsonObject.put(key, (map.get(key)));
        }
        return jsonObject;
    }
}
