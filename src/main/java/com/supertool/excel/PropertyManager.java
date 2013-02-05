package com.supertool.excel;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import org.apache.log4j.Logger;

public class PropertyManager {
	static Logger logger = Logger.getLogger(PropertyManager.class);
    private Properties prop=null;
    private Map<String,List<String>> keyMaps=null;
    
    private final String SHEET_NUMBERS="sheet.number";
    private final String AREA_NUMBERS="area.number";
    private final String AREA="area";
    private final String KEYS="keys";
    private final String START_POINT="startpoint";
    private final String END_POINT="endpoint";
    private final String KEYS_SPLIT=";";
    private final String POINT_SPLIT=",";
    
    public boolean openFile(InputStream is){
    	boolean result=true;
    	prop = new Properties();
		try {
			prop.load(is);
		} catch (IOException e) {
			logger.error("Read properties file error!\n");
	        result=false;
		}
    	return result;
    }
    
    public boolean openFile(String path){
    	boolean result=true;
    	prop = new Properties();
    	try {
    		InputStream stream = this.getClass().getClassLoader().getResourceAsStream(path);
			prop.load(stream);
		} catch (FileNotFoundException e) {
            logger.error("Properties file not exits!\n");
            result=false;
		} catch (IOException e) {
            logger.error("Read properties file error!\n");
            result=false;
		}
    	return result;
    }
    
    /**
     * 返回sheet的数量
     * @return 返回-1 表示只有sheet0为模板返回其他数据表示0-指定值为模版
     */
    public int getSheetNumbers(){
    	String numbers=this.prop.getProperty(this.SHEET_NUMBERS).trim();
    	return Integer.parseInt(numbers);
    }
    
    /**
     * 返回 获得所有的ExCell 
     * @return
     */
    public List<ExCell> getCells(){
    	List<String> keys=this.keyMaps.get(AREA+0);
    	List<ExCell> cells=new ArrayList<ExCell>();
    	String key=null;
    	ExCell excell=null;
    	for(int i=0;i<keys.size();i++){
    		key=keys.get(i);
    		String pointStr=this.prop.getProperty(key);
    		excell=this.getCellFromProperty(pointStr);
    		excell.setKey(key);
    		cells.add(excell);
    	}
    	return cells;
    }
    
    /**
     * 获得keys
     * @return 返回area0的key，area1的key。
     */
    public Map<String,List<String>> getKeys(){
    	keyMaps=new HashMap<String,List<String>>();
    	List<String> areas=this.getAreas();
    	for(int i=0;i<areas.size();i++){
    		String area=areas.get(i);
    		String keysStr=this.prop.getProperty(area+"."+this.KEYS).trim();
    		keyMaps.put(area,this.getKeyFromList(keysStr));
    	}
    	return keyMaps;
    }
    
    /**
     * 获得数据域的起始点和结束点
     * @return
     */
    public Map<String,ExCell> getData(){
    	Map<String,ExCell> dataMap= new HashMap<String,ExCell>();
    	String startPointStr=this.prop.getProperty(this.AREA+"1."+this.START_POINT);
    	String endPointStr=this.prop.getProperty(this.AREA+"1."+this.END_POINT);
    	dataMap.put(START_POINT,this.getCellFromProperty(startPointStr));
    	dataMap.put(END_POINT, this.getCellFromProperty(endPointStr));
    	return dataMap;
    }
    
    /**
     * 获得所有域
     * @return 返回所有域。
     */
    private List<String> getAreas(){
    	List<String> areas=new ArrayList<String>();
    	int areaNumber=Integer.parseInt(prop.getProperty(AREA_NUMBERS).toString());
    	for(int i=0;i<areaNumber;i++){
    		areas.add(this.AREA+i);
    	}
    	return areas;
    }
    
    /**
     * 获得域的keys。
     * @param keysStr keys字符串。
     * @return 返回所有域。
     */
    private List<String> getKeyFromList(String keysStr){
    	 List<String> keysList=new ArrayList<String>();
         String[] keys=keysStr.split(KEYS_SPLIT);
         for(int i=0;i<keys.length;i++){
        	 keysList.add(keys[i]);
         }  
         return keysList;
    }
    
    /**
     * 获得单元格的坐标
     * @param point 坐标的字符串
     * @return 返回ExCell 包括cell的坐标和key
     */
    private ExCell getCellFromProperty(String point){
    	ExCell excell=new ExCell();
    	String[] pointProp=point.split(POINT_SPLIT);
    	excell.setX(Integer.parseInt(pointProp[0]));
    	excell.setY(Integer.parseInt(pointProp[1]));
    	return excell;
    }
}
