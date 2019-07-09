package com.study.util;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import com.thinkive.base.util.StringHelper;

/**
 * 
 * @描述: 题目对象实体
 * @版权: Copyright (c) 2019 
 * @公司: 思迪科技 
 * @作者: 严磊
 * @版本: 1.0 
 * @创建日期: 2019年7月8日 
 * @创建时间: 下午8:14:48
 */
public class SubjectEntity
{
    public SubjectEntity(){
        this.options = new ArrayList<List<String>>();
    }
    //题目信息
    private String content; //问题题目
    private String type; //题目类型；0:单选，1、多选，2：问答
    private String rank; //排序值
    private String is_must; //是否为必答题；0:非必答题，1、必答题
    private String is_myd; //是否满意度题目
    //选项信息
    List<List<String>> options; //选项排序    选项分值    是否正确选项（0:否/1:是） 选项内容
    public String getContent()
    {
        return content;
    }
    public void setContent(String content)
    {
        if(StringHelper.isBlank(content))
        {
            new RuntimeException("题干内容为空");
        }
        else
        {
            this.content = content;
        }
    }
    public String getType()
    {
        return type;
    }
    public void setType(String type)
    {
        if(StringHelper.isBlank(type))
        {
            this.type = "0";
        }
        else
        {
            this.type = type.trim();
        }
    }
    public String getRank()
    {
        return rank;
    }
    public void setRank(String rank)
    {
        this.rank = rank.trim();
    }
    public String getIs_must()
    {
        return is_must;
    }
    public void setIs_must(String is_must)
    {
        if(StringHelper.isBlank(is_must))
        {
            this.is_must = "1";
        }
        else
        {
            this.is_must = is_must.trim();
        }
    }
    public String getIs_myd()
    {
        return is_myd;
    }
    public void setIs_myd(String is_myd)
    {
        if(StringHelper.isBlank(is_myd))
        {
            this.is_myd = "0";
        }
        else
        {
            this.is_myd = is_myd.trim();
        }
    }
    
    public List<List<String>> getOptions()
    {
        //对选项进行重排序
        Collections.sort(this.options, new Comparator<List<String>>() {
            @Override
            public int compare(List<String> L1, List<String> L2) {
                int sort1 = Integer.parseInt(L1.get(0).trim());
                int sort2 = Integer.parseInt(L1.get(0).trim());
                return sort1-sort2;
            }
        });
        for (int i = 0; i < this.options.size(); i++)
        {
            List<String> list = this.options.get(i);
            list.set(0, i+1+"");
        }
        return options;
    }
    
    public void addOptions(List<String> option)
    {
        this.options.add(option);
    }
    @Override
    public String toString()
    {
        return "SubjectEntity [content=" + content + ", type=" + type + ", rank=" + rank + ", is_must=" + is_must
                + ", is_myd=" + is_myd + ", options=" + options.toString() + "]";
    }
    
    
}
