package com.holiday.umengparse.model;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by yuan on 2018/3/29.
 *
 * cell：代表一个逻辑意义的cell
 */
public class MergeCell {
    public String content;
    public CellRangeAddress mergeRegion;
    public List<MergeCell> children = new ArrayList<>();
}
