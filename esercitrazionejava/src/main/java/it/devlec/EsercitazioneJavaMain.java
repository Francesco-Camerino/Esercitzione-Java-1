package it.devlec;

import it.devlec.excel.EsempioExcel;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class EsercitazioneJavaMain {
    private static final Logger logger =  LogManager.getLogger(EsercitazioneJavaMain.class);
    public static void main(String[] args){
        /*logger.trace("Hello from Log4j 2");
        logger.debug("Hello from Log4j 2");
        logger.info("Hello from Log4j 2");
        logger.warn("Hello from Log4j 2");
        logger.error("Hello from Log4j 2");
        logger.fatal("Hello from Log4j 2");
        EsempioLog esempioLog = new EsempioLog();
        esempioLog.stampaAltriLog();*/
        /*EsempioCSV esempioCSV = new EsempioCSV();
        MioCSV mioCSV = new MioCSV();
        esempioCSV.leggiCSV();
        esempioCSV.scriviCSV();
        mioCSV.leggiIlMioCSV();
        mioCSV.scriviIlmioCSV();
        mioCSV.leggiIlMioCSV();*/
        EsempioExcel esempioExcel = new EsempioExcel();
        esempioExcel.testLetturaExcel();
        esempioExcel.scriviIlMioFileExcel();
        esempioExcel.leggiDaCSV();
       /* EsempioPDF esempioPDF = new EsempioPDF();
        esempioPDF.creaMioPdf();
        EsempioJSON esempioJSON = new EsempioJSON();
        esempioJSON.esempioJSONOggetto();
        esempioJSON.esempioJSONArray();*/
    }
}
