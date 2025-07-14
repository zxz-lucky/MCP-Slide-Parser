package com.zxz.mcp.mcpslideparser.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 表示表格形状
 */

public class Table {

    private int rows;
    private int columns;
    private List<Cell> cells = new ArrayList<>();
    private Style style;

    public static class Cell {
        private int row;
        private int column;
        private String text;
        private Style style;

        // Getters and Setters
        public int getRow() {
            return row;
        }

        public void setRow(int row) {
            this.row = row;
        }

        public int getColumn() {
            return column;
        }

        public void setColumn(int column) {
            this.column = column;
        }

        public String getText() {
            return text;
        }

        public void setText(String text) {
            this.text = text;
        }

        public Style getStyle() {
            return style;
        }

        public void setStyle(Style style) {
            this.style = style;
        }
    }

    // Getters and Setters
    public int getRows() {
        return rows;
    }

    public void setRows(int rows) {
        this.rows = rows;
    }

    public int getColumns() {
        return columns;
    }

    public void setColumns(int columns) {
        this.columns = columns;
    }

    public List<Cell> getCells() {
        return cells;
    }

    public void addCell(Cell cell) {
        this.cells.add(cell);
    }

    public Style getStyle() {
        return style;
    }

    public void setStyle(Style style) {
        this.style = style;
    }
}