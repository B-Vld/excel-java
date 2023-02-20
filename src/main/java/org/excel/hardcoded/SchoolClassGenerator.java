package org.excel.hardcoded;

import com.github.javafaker.Faker;
import lombok.NoArgsConstructor;
import org.excel.dao.SchoolClass;
import org.excel.dao.Student;
import org.excel.dao.Subject;

import java.util.*;
import java.util.function.Supplier;
import java.util.stream.Stream;

@NoArgsConstructor
public final class SchoolClassGenerator {

    private static final Faker FAKER = new Faker();

    public static List<SchoolClass> generateClasses(int nClasses, int nStudents) {
        return Stream
                .generate(generateClass(nStudents))
                .limit(nClasses)
                .toList();
    }

    public static List<Student> generateStudents(int count) {
        return Stream
                .generate(generateStudent())
                .limit(count)
                .toList();
    }

    public static Supplier<SchoolClass> generateClass(int nStudents) {
        return () -> SchoolClass.builder()
                .id(UUID.randomUUID())
                .subject(FAKER.superhero().power())
                .students(generateStudents(nStudents))
                .build();
    }


    public static Supplier<Student> generateStudent() {
        return () -> Student.builder()
                .id(UUID.randomUUID())
                .age(FAKER.number().numberBetween(18, 80))
                .firstName(FAKER.funnyName().name())
                .lastName(FAKER.funnyName().name())
                .grades(generateGrades())
                .build();
    }

    public static Map<Subject, List<Integer>> generateGrades() {
        var grades = new HashMap<Subject, List<Integer>>();
        Arrays.stream(Subject.values()).forEach(subject -> {
            var numberOfGradesPerSubject = FAKER.number().numberBetween(3, 7);
            var gradeList = Stream
                    .generate(generateGrade())
                    .limit(numberOfGradesPerSubject)
                    .toList();
            grades.put(subject, gradeList);
        });
        return grades;
    }

    public static Supplier<Integer> generateGrade() {
        return () -> FAKER.number().numberBetween(1, 11);
    }


}
