package org.excel.dao;

import lombok.Builder;
import lombok.Data;

import java.util.List;
import java.util.UUID;

@Data
@Builder
public class SchoolClass {

    private UUID id;
    private String subject;
    private List<Student> students;

}
