package com.study.util;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import com.thinkive.base.util.StringHelper;

/**
 * 
 * @����: ��Ŀ����ʵ��
 * @��Ȩ: Copyright (c) 2019 
 * @��˾: ˼�ϿƼ� 
 * @����: ����
 * @�汾: 1.0 
 * @��������: 2019��7��8�� 
 * @����ʱ��: ����8:14:48
 */
public class SubjectEntity
{
    public SubjectEntity(){
        this.options = new ArrayList<List<String>>();
    }
    //��Ŀ��Ϣ
    private String content; //������Ŀ
    private String type; //��Ŀ���ͣ�0:��ѡ��1����ѡ��2���ʴ�
    private String rank; //����ֵ
    private String is_must; //�Ƿ�Ϊ�ش��⣻0:�Ǳش��⣬1���ش���
    private String is_myd; //�Ƿ��������Ŀ
    //ѡ����Ϣ
    List<List<String>> options; //ѡ������    ѡ���ֵ    �Ƿ���ȷѡ�0:��/1:�ǣ� ѡ������
    public String getContent()
    {
        return content;
    }
    public void setContent(String content)
    {
        if(StringHelper.isBlank(content))
        {
            new RuntimeException("�������Ϊ��");
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
        //��ѡ�����������
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
