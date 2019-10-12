package com.alibaba.excel.analysis.v03.handlers;

import com.alibaba.excel.analysis.v03.AbstractXlsRecordHandler;
import com.alibaba.excel.metadata.CellData;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;

import java.math.BigDecimal;
import java.math.RoundingMode;

/**
 * Record handler
 *
 * @author Dan Zheng
 */
public class NumberRecordHandler extends AbstractXlsRecordHandler {
    private FormatTrackingHSSFListener formatListener;

    /**
     * precision
     */
    private static final double EPS = 1e-10;

    public NumberRecordHandler(FormatTrackingHSSFListener formatListener) {
        this.formatListener = formatListener;
    }

    @Override
    public boolean support(Record record) {
        return NumberRecord.sid == record.getSid();
    }

    @Override
    public void processRecord(Record record) {
        NumberRecord numrec = (NumberRecord)record;
        this.row = numrec.getRow();
        this.column = numrec.getColumn();

        BigDecimal value = BigDecimal.valueOf(numrec.getValue());
        if (isIntegerForDouble(numrec.getValue())) {
            // remove .0 for integer number
            value = value.setScale(0, RoundingMode.DOWN);
        }
        this.cellData = new CellData(value);
        this.cellData.setDataFormat(formatListener.getFormatIndex(numrec));
        this.cellData.setDataFormatString(formatListener.getFormatString(numrec));
    }

    /**
     * check double number whether is an integer.
     * @param num double number
     * @return true - integer number; false - number with point
     */
    public static boolean isIntegerForDouble(double num) {
        return num - Math.floor(num) < EPS;
    }

    @Override
    public void init() {

    }

    @Override
    public int getOrder() {
        return 0;
    }
}
