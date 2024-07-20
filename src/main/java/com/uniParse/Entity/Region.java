package com.uniParse.Entity;

import lombok.Data;

@Data
public class Region {
    private String ref;
    private String name;

    public Region() {
    }

    public Region(String ref, String name) {
        this.ref = ref;
        this.name = name;
    }

    @Override
    public String toString() {
        return "TableCell{" +
                "href='" + ref + '\'' +
                ", name='" + name + '\'' +
                '}';
    }

    public String getName() {
        return name;
    }

    public String getRef() {
        return ref;
    }
}
