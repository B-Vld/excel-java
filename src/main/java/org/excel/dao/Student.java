package org.excel.dao;

import lombok.Builder;
import lombok.Data;

import java.util.List;
import java.util.Map;
import java.util.UUID;

@Data
@Builder
public class Student {

    private UUID id;
    private String firstName;
    private String lastName;
    private int age;
    private Map<Subject, List<Integer>> grades;

}
