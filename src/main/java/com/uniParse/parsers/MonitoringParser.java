package com.uniParse.parsers;

import com.uniParse.Entity.Region;
import com.uniParse.Entity.University;
import org.jsoup.Connection;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class MonitoringParser {
    Elements title;

    public List<Region> getRegionRefList(Connection connection){
        List<Region> regions = new ArrayList<>();
        Document page = getConnection(connection);

        title = page.getElementsByAttributeValue("id","tregion").select("table > tbody > tr > td > p");

        for (org.jsoup.nodes.Element element : title) {
            if(element.text().matches("^.*федеральный округ$")) continue;
            Region region =
                    new Region(element.getElementsByTag("a")
                            .attr("href"), element.text());
            regions.add(region);
        }
        return regions;
    }

    public List<University> getUniversityRefList(Connection connection){
        List<University> universities = new ArrayList<>();
        Document page = getConnection(connection);

        title = page.getElementsByAttributeValue("class", "an").select("table > tbody > tr");
        for (org.jsoup.nodes.Element element : title) {
            Elements hrefElement = element.select("tr > td > a");
            universities.add(new University(hrefElement.attr("href"), hrefElement.text()));
        }
        return universities;
    }

    public Elements getResultTable(Connection connection){
        Document page = getConnection(connection);

        //получаем одну таблицу
        return title = page.getElementsByAttributeValue("id", "result");
    }

    public Elements getListOfTablesWithAttributeNapde(Connection connection){
        Document page = getConnection(connection);

        //получаем несколько таблиц
        return title = page.getElementsByAttributeValue("class", "napde");
    }

    public Elements getTableAnalisReg(Connection connection){
        Document page = getConnection(connection);

        //получаем одну таблицу
        return title = page.getElementsByAttributeValue("id", "analis_reg");
    }

    public Elements getTableKontByOtr(Connection connection){
        Document page = getConnection(connection);

        //получаем одну таблицу
        return title= page.getElementsByAttributeValue("id", "kont_by_otr");
    }

    public Elements getTableAnalisDop(Connection connection){
        Document page = getConnection(connection);

        //получаем одну таблицу
        return title = page.getElementsByAttributeValue("id", "analis_dop");
    }


    private Document getConnection(Connection connection){
        Document page;
        try {
            page = connection.get();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return page;
    }
}
