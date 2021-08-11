package org.cp;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

public class SimpleReadListener extends AnalysisEventListener<Object>
{
	private List<Map<String,String>> list=new ArrayList<Map<String,String>>();
	private Map<Integer,String> headMap;

	@Override
	public void invokeHeadMap(Map<Integer,String> headMap,AnalysisContext context)
	{
		this.headMap=headMap;
	}

	@SuppressWarnings("unchecked")
	@Override
	public void invoke(Object object,AnalysisContext context)
	{
		Map<Integer,String> data=(Map<Integer,String>)object;
		Map<String,String> v=new LinkedHashMap<>(data.size());
		for(Entry<Integer,String> e:data.entrySet())
		{
			v.put(headMap.get(e.getKey()),e.getValue());
		}
		list.add(v);
	}

	@Override
	public void doAfterAllAnalysed(AnalysisContext context)
	{
		// do nothing
	}

	public List<Map<String,String>> getList()
	{
		return list;
	}

	public Map<Integer,String> getHeadMap()
	{
		return headMap;
	}
}