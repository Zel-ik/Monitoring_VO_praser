package com.uniParse;

import com.uniParse.Entity.Region;
import com.uniParse.Entity.University;
import com.uniParse.converters.ExcelConverter;
import com.uniParse.parsers.MonitoringParser;
import org.jsoup.Jsoup;
import org.jsoup.select.Elements;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {


    public static void main(String[] args) {
        ExcelConverter excelConverter = new ExcelConverter();
        MonitoringParser monitoringParser = new MonitoringParser();
        Map<String, List<University>> allUniversityHashMap = new HashMap<>();

        String connection = "https://monitoring.miccedu.ru/";

        List<Region> regionRefs = monitoringParser
                .getRegionRefList(Jsoup.connect(connection + "?m=vpo&year=2020"));

        List<University> universityList;
        for (Region region : regionRefs) {
            universityList = monitoringParser
                    .getUniversityRefList(Jsoup.connect(connection + region.getRef()));
            allUniversityHashMap.put(region.getName(), universityList);
        }

        for(int i = 0; i < allUniversityHashMap.size(); i++){
            excelConverter.createWorkbook(regionRefs.get(i).getName(), allUniversityHashMap.get(regionRefs.get(i).getName()));
        }

        Elements resultTable = monitoringParser.getResultTable(Jsoup.connect(connection + "iam/2019/_vpo/" +
                allUniversityHashMap.get(regionRefs.get(0).getName())
                        .get(0).getRef())
        );

        Elements tablesWithAttributeNapde = monitoringParser.getListOfTablesWithAttributeNapde(
                Jsoup.connect(connection + "iam/2019/_vpo/" +
                        allUniversityHashMap.get(regionRefs.get(0).getName())
                                .get(0).getRef())
        );

        Elements tableWithAttributeKontByOtr = monitoringParser.getTableKontByOtr(
                Jsoup.connect(connection + "iam/2019/_vpo/" +
                        allUniversityHashMap.get(regionRefs.get(0).getName())
                                .get(0).getRef()));

        Elements tableAnalisReq = monitoringParser.getTableAnalisReg(Jsoup.connect(connection + "iam/2019/_vpo/" +
                allUniversityHashMap.get(regionRefs.get(0).getName())
                        .get(0).getRef())
        );

        Elements tableAnalisDop = monitoringParser.getTableAnalisDop(Jsoup.connect(connection + "iam/2019/_vpo/" +
                allUniversityHashMap.get(regionRefs.get(0).getName())
                        .get(0).getRef())
        );

        for(int i =0; i < regionRefs.size(); i++){
            for(int b = 0; b < allUniversityHashMap.get(regionRefs.get(i).getName()).size(); b++){
                excelConverter.writeTablesIntoExcel(regionRefs.get(i).getName(),
                        String.valueOf(b),
                        resultTable,
                        tablesWithAttributeNapde,
                        tableWithAttributeKontByOtr,
                        tableAnalisReq,
                        tableAnalisDop);
            }
        }

    }
}