/*
 * Copyright 2018 NingWei (ningww1@126.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * </p>
 */
package com.excel.poi;

import com.excel.poi.common.Constant;
import com.excel.poi.entity.ExcelEntity;
import com.excel.poi.excel.ExcelReader;
import com.excel.poi.excel.ExcelWriter;
import com.excel.poi.exception.ExcelBootException;
import com.excel.poi.factory.ExcelMappingFactory;
import com.excel.poi.function.ExportFunction;
import com.excel.poi.function.ImportFunction;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import javax.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.xml.sax.SAXException;

/**
 * @author NingWei
 */
@Slf4j
public class ExcelBoot {
    private HttpServletResponse httpServletResponse;
    private OutputStream outputStream;
    private InputStream inputStream;
    private String fileName;
    private Class excelClass;
    private Integer pageSize;
    private Integer rowAccessWindowSize;
    private Integer recordCountPerSheet;
    private Boolean openAutoColumWidth;


    protected ExcelBoot(InputStream inputStream, Class excelClass) {
        this(null, null, inputStream, null, excelClass, null, null, null, null);
    }


    protected ExcelBoot(OutputStream outputStream, String fileName, Class excelClass) {
        this(null, outputStream, null, fileName, excelClass, Constant.DEFAULT_PAGE_SIZE, Constant.DEFAULT_ROW_ACCESS_WINDOW_SIZE, Constant.DEFAULT_RECORD_COUNT_PEER_SHEET, Constant.OPEN_AUTO_COLUM_WIDTH);
    }


    protected ExcelBoot(HttpServletResponse response, String fileName, Class excelClass) {
        this(response, null, null, fileName, excelClass, Constant.DEFAULT_PAGE_SIZE, Constant.DEFAULT_ROW_ACCESS_WINDOW_SIZE, Constant.DEFAULT_RECORD_COUNT_PEER_SHEET, Constant.OPEN_AUTO_COLUM_WIDTH);
    }


    protected ExcelBoot(HttpServletResponse response, OutputStream outputStream, InputStream inputStream
            , String fileName, Class excelClass, Integer pageSize, Integer rowAccessWindowSize, Integer recordCountPerSheet, Boolean openAutoColumWidth) {
        this.httpServletResponse = response;
        this.outputStream = outputStream;
        this.inputStream = inputStream;
        this.fileName = fileName;
        this.excelClass = excelClass;
        this.pageSize = pageSize;
        this.rowAccessWindowSize = rowAccessWindowSize;
        this.recordCountPerSheet = recordCountPerSheet;
        this.openAutoColumWidth = openAutoColumWidth;
    }


    public static ExcelBoot ExportBuilder(HttpServletResponse httpServletResponse, String fileName, Class clazz) {
        return new ExcelBoot(httpServletResponse, fileName, clazz);
    }


    public static ExcelBoot ExportBuilder(OutputStream outputStream, String fileName, Class clazz) {
        return new ExcelBoot(outputStream, fileName, clazz);
    }


    public static ExcelBoot ExportBuilder(HttpServletResponse response, String fileName, Class excelClass,
                                          Integer pageSize, Integer rowAccessWindowSize, Integer recordCountPerSheet, Boolean openAutoColumWidth) {
        return new ExcelBoot(response, null, null
                , fileName, excelClass, pageSize, rowAccessWindowSize, recordCountPerSheet, openAutoColumWidth);
    }


    public static ExcelBoot ExportBuilder(OutputStream outputStream, String fileName, Class excelClass, Integer pageSize
            , Integer rowAccessWindowSize, Integer recordCountPerSheet, Boolean openAutoColumWidth) {
        return new ExcelBoot(null, outputStream, null
                , fileName, excelClass, pageSize, rowAccessWindowSize, recordCountPerSheet, openAutoColumWidth);
    }


    public static ExcelBoot ImportBuilder(InputStream inputStreamm, Class clazz) {
        return new ExcelBoot(inputStreamm, clazz);
    }


    public <R, T> void exportResponse(R param, ExportFunction<R, T> exportFunction) {
        SXSSFWorkbook sxssfWorkbook = null;
        try {
            try {
                verifyResponse();
                sxssfWorkbook = commonSingleSheet(param, exportFunction);
                download(sxssfWorkbook, httpServletResponse, URLEncoder.encode(fileName + ".xlsx", "UTF-8"));
            } finally {
                if (sxssfWorkbook != null) {
                    sxssfWorkbook.close();
                }
                if (httpServletResponse != null && httpServletResponse.getOutputStream() != null) {
                    httpServletResponse.getOutputStream().close();
                }
            }
        } catch (Exception e) {
            throw new ExcelBootException(e);
        }
    }


    public <R, T> void exportStream(R param, ExportFunction<R, T> exportFunction) {
        OutputStream outputStream = null;
        try {
            try {
                outputStream = generateStream(param, exportFunction);
                write(outputStream);
            } finally {
                if (outputStream != null) {
                    outputStream.close();
                }
            }
        } catch (Exception e) {
            throw new ExcelBootException(e);
        }
    }


    public <R, T> OutputStream generateStream(R param, ExportFunction<R, T> exportFunction) throws IOException {
        SXSSFWorkbook sxssfWorkbook = null;
        try {
            verifyStream();
            sxssfWorkbook = commonSingleSheet(param, exportFunction);
            sxssfWorkbook.write(outputStream);
            return outputStream;
        } catch (Exception e) {
            log.error("生成Excel发生异常! 异常信息:", e);
            if (sxssfWorkbook != null) {
                sxssfWorkbook.close();
            }
            throw new ExcelBootException(e);
        }
    }


    public <R, T> void exportMultiSheetResponse(R param, ExportFunction<R, T> exportFunction) {
        SXSSFWorkbook sxssfWorkbook = null;
        try {
            try {
                verifyResponse();
                sxssfWorkbook = commonMultiSheet(param, exportFunction);
                download(sxssfWorkbook, httpServletResponse, URLEncoder.encode(fileName + ".xlsx", "UTF-8"));
            } finally {
                if (sxssfWorkbook != null) {
                    sxssfWorkbook.close();
                }
            }
        } catch (Exception e) {
            throw new ExcelBootException(e);
        }
    }



    public <R, T> void exportMultiSheetStream(R param, ExportFunction<R, T> exportFunction) {
        OutputStream outputStream = null;
        try {
            try {
                outputStream = generateMultiSheetStream(param, exportFunction);
                write(outputStream);
            } finally {
                if (outputStream != null) {
                    outputStream.close();
                }
            }
        } catch (Exception e) {
            throw new ExcelBootException(e);
        }
    }



    public <R, T> OutputStream generateMultiSheetStream(R param, ExportFunction<R, T> exportFunction) throws IOException {
        SXSSFWorkbook sxssfWorkbook = null;
        try {
            verifyStream();
            sxssfWorkbook = commonMultiSheet(param, exportFunction);
            sxssfWorkbook.write(outputStream);
            return outputStream;
        } catch (Exception e) {
            log.error("分Sheet生成Excel发生异常! 异常信息:", e);
            if (sxssfWorkbook != null) {
                sxssfWorkbook.close();
            }
            throw new ExcelBootException(e);
        }
    }


    public void exportTemplate() {
        SXSSFWorkbook sxssfWorkbook = null;
        try {
            try {
                verifyResponse();
                verifyParams();
                ExcelEntity excelMapping = ExcelMappingFactory.loadExportExcelClass(excelClass, fileName);
                ExcelWriter excelWriter = new ExcelWriter(excelMapping, pageSize, rowAccessWindowSize, recordCountPerSheet, openAutoColumWidth);
                sxssfWorkbook = excelWriter.generateTemplateWorkbook();
                download(sxssfWorkbook, httpServletResponse, URLEncoder.encode(fileName + ".xlsx", "UTF-8"));
            } finally {
                if (sxssfWorkbook != null) {
                    sxssfWorkbook.close();
                }
                if (httpServletResponse != null && httpServletResponse.getOutputStream() != null) {
                    httpServletResponse.getOutputStream().close();
                }
            }
        } catch (Exception e) {
            throw new ExcelBootException(e);
        }
    }

    public void importExcel(ImportFunction importFunction) {
        try {
            if (importFunction == null) {
                throw new ExcelBootException("excelReadHandler参数为空!");
            }
            if (inputStream == null) {
                throw new ExcelBootException("inputStream参数为空!");
            }

            ExcelEntity excelMapping = ExcelMappingFactory.loadImportExcelClass(excelClass);
            ExcelReader excelReader = new ExcelReader(excelClass, excelMapping, importFunction);
            excelReader.process(inputStream);
        } catch (Exception e) {
            throw new ExcelBootException(e);
        }

    }

    private <R, T> SXSSFWorkbook commonSingleSheet(R param, ExportFunction<R, T> exportFunction) throws Exception {
        verifyParams();
        ExcelEntity excelMapping = ExcelMappingFactory.loadExportExcelClass(excelClass, fileName);
        ExcelWriter excelWriter = new ExcelWriter(excelMapping, pageSize, rowAccessWindowSize, recordCountPerSheet, openAutoColumWidth);
        return excelWriter.generateWorkbook(param, exportFunction);
    }

    private <R, T> SXSSFWorkbook commonMultiSheet(R param, ExportFunction<R, T> exportFunction) throws Exception {
        verifyParams();
        ExcelEntity excelMapping = ExcelMappingFactory.loadExportExcelClass(excelClass, fileName);
        ExcelWriter excelWriter = new ExcelWriter(excelMapping, pageSize, rowAccessWindowSize, recordCountPerSheet, openAutoColumWidth);
        return excelWriter.generateMultiSheetWorkbook(param, exportFunction);
    }


    private void write(OutputStream out) throws IOException {
        if (null != out) {
            out.flush();
        }
    }



    private void download(SXSSFWorkbook wb, HttpServletResponse response, String filename) throws IOException {
        OutputStream out = response.getOutputStream();
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-disposition",
                String.format("attachment; filename=%s", filename));
        if (null != out) {
            wb.write(out);
            out.flush();
        }
    }

    private void verifyResponse() {
        if (httpServletResponse == null) {
            throw new ExcelBootException("httpServletResponse参数为空!");
        }
    }

    private void verifyStream() {
        if (outputStream == null) {
            throw new ExcelBootException("outputStream参数为空!");
        }
    }

    private void verifyParams() {
        if (excelClass == null) {
            throw new ExcelBootException("excelClass参数为空!");
        }
        if (fileName == null) {
            throw new ExcelBootException("fileName参数为空!");
        }
    }

}