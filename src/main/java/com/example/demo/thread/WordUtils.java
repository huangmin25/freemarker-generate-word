package com.example.demo.thread;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import sun.misc.BASE64Encoder;

/**
 * 类说明：
 * 
 * @version 2.0
 *          <p>
 * 
 *          <pre>
 *   文件名:    demo.java  <br>
 *   创建人:    admin <br>
 *   时间:     2015年7月24日<br>
 * </pre>
 * 
 *          </p>
 **/

public class WordUtils {

	private Configuration configuration = null;

	public WordUtils() {
		configuration = new Configuration();
		configuration.setDefaultEncoding("UTF-8");
	}

	public static void main(String[] args) {
		WordUtils dm = new WordUtils();
		dm.createWord();
	}

	public void createWord() {
		Map<String, Object> dataMap = new HashMap<String, Object>();
		getDataForDomox(dataMap);
		configuration.setClassForTemplateLoading(this.getClass(), "/static"); // FTL文件所存在的位置
		Template t = null;
		try {
			t = configuration.getTemplate("demox.ftl"); // 文件名
		} catch (IOException e) {
			e.printStackTrace();
		}
		File outFile = new File("D:/demox.doc");
		Writer out = null;
		try {
			out = new BufferedWriter(new OutputStreamWriter(
					new FileOutputStream(outFile)));
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		}

		try {
			t.process(dataMap, out);
		} catch (TemplateException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void getData(Map<String, Object> dataMap) {
		dataMap.put("Title", "文档标题");
		dataMap.put("date",
				new SimpleDateFormat("yyyy-MM-dd").format(new Date()));

		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		for (int i = 0; i < 10; i++) {
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("Number", i);
			map.put("content", "文本内容" + i);
			list.add(map);
		}
		dataMap.put("newList", list);
	}
	
	private void getDataForDomox(Map<String, Object> dataMap) {
		dataMap.put("userName", "文档标题");
		dataMap.put("currDate",
				new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
		dataMap.put("content", "文档标题文档标题文档标题文档标题文档标题文档标题文档标题文档标题文档标题文档标题文档标题文档标题文档标题文档标题文档标题文档标题文档标题文档标题");

		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		
		String img = null;
        InputStream in;
        byte[] picdata = null;
        try {
            in = new FileInputStream("D:\\02.jpg");
            picdata = new byte[in.available()];
            in.read(picdata);
            in.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        BASE64Encoder encoder = new BASE64Encoder();
        img = encoder.encode(picdata);
		for (int i = 0; i < 10; i++) {
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("title", img);
			map.put("content", "文本内容" + i);
			map.put("author", "作者" + i);
			list.add(map);
		}
		dataMap.put("newList", list);
	}
}