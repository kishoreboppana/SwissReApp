package com.example;
// Java 11 code to address all the use cases listed

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.*;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.*;
import java.util.stream.Collectors;

class Employee {
    int id;
    String name;
    String city;
    String state;
    String category;
    Integer managerId;
    double salary;
    LocalDate doj;

    public Employee(int id, String name, String city, String state, String category, Integer managerId, double salary, LocalDate doj) {
        this.id = id;
        this.name = name;
        this.city = city;
        this.state = state;
        this.category = category;
        this.managerId = managerId;
        this.salary = salary;
        this.doj = doj;
    }
}

public class EmployeeExcelProcessor {

    private static final String[] HEADER = {"id", "name", "city", "state", "category", "manager_id", "salary", "DOJ"};
    private static final DateTimeFormatter formatter = DateTimeFormatter.ofPattern("d-MMM-yyyy");

    public static void main(String[] args) throws Exception {
        List<Employee> employees = generateDummyEmployees();
        writeEmployeesToExcel(employees, "employees.xlsx");
        List<Employee> loadedEmployees = readEmployeesFromExcel("employees.xlsx");

        List<Employee> gratuityEligible = findGratuityEligibleEmployees(loadedEmployees);
        System.out.println("\nGratuity Eligible Employees:");
        gratuityEligible.forEach(e -> System.out.println(e.name));

        List<Employee> higherPaidThanManager = findEmployeesWithHigherSalaryThanManager(loadedEmployees);
        System.out.println("\nEmployees earning more than their managers:");
        higherPaidThanManager.forEach(e -> System.out.println(e.name));

        buildOrgHierarchyJson(loadedEmployees, "employee_hierarchy.json");
    }

    private static List<Employee> generateDummyEmployees() {
        List<Employee> list = new ArrayList<>();

        list.add(new Employee(789, "Rama", "chennai", "Tamilnadu", "Director", null, 150000, LocalDate.parse("25-Oct-2022", formatter)));
        list.add(new Employee(456, "Shivam", "bangalore", "karnataka", "manager", 789, 75000, LocalDate.parse("5-Jul-2022", formatter)));
        list.add(new Employee(123, "Ravi", "hyderabad", "Telangana", "employee", 456, 45000, LocalDate.parse("4-Jun-2023", formatter)));
        list.add(new Employee(1011, "Krishna", "hyderabad", "telangana", "employee", 456, 50000, LocalDate.parse("7-Mar-2021", formatter)));
        list.add(new Employee(1213, "Sreekanth", "mumbai", "Maharastra", "employee", 789, 60000, LocalDate.parse("8-Aug-2019", formatter)));
        list.add(new Employee(1415, "Manoj", "mangalore", "karnataka", "employee", 456, 95000, LocalDate.parse("9-Jun-2018", formatter)));
        list.add(new Employee(2000, "Sanjay0", "kochi", "Kerala", "manager", 789, 103022, LocalDate.parse("24-Aug-2022", formatter)));
        list.add(new Employee(2001, "Rahul1", "noida", "UP", "manager", 789, 40691, LocalDate.parse("30-Sep-2023", formatter)));
        list.add(new Employee(2002, "Neha2", "noida", "UP", "manager", 789, 90044, LocalDate.parse("29-Sep-2016", formatter)));
        list.add(new Employee(2003, "Neha3", "noida", "UP", "employee", 456, 96986, LocalDate.parse("4-Jul-2018", formatter)));
        list.add(new Employee(2004, "Arjun4", "pune", "Maharashtra", "employee", 456, 79918, LocalDate.parse("31-Oct-2016", formatter)));
        list.add(new Employee(2005, "Neha5", "pune", "Maharashtra", "manager", 789, 46032, LocalDate.parse("2-Dec-2016", formatter)));
        list.add(new Employee(2006, "Divya6", "indore", "MP", "employee", 456, 114576, LocalDate.parse("2-Dec-2016", formatter)));
        list.add(new Employee(2007, "Sanjay7", "kochi", "Kerala", "manager", 789, 50369, LocalDate.parse("7-Feb-2020", formatter)));
        list.add(new Employee(2008, "Sanjay8", "delhi", "Delhi", "manager", 789, 113213, LocalDate.parse("3-Jan-2018", formatter)));
        list.add(new Employee(2009, "Rahul9", "kochi", "Kerala", "employee", 456, 47873, LocalDate.parse("17-Nov-2018", formatter)));
        list.add(new Employee(2010, "Rahul10", "noida", "UP", "manager", 789, 96274, LocalDate.parse("10-Jun-2016", formatter)));
        list.add(new Employee(2011, "Amit11", "indore", "MP", "manager", 789, 106637, LocalDate.parse("16-Mar-2023", formatter)));
        list.add(new Employee(2012, "Sanjay12", "noida", "UP", "manager", 789, 60379, LocalDate.parse("28-Jan-2024", formatter)));
        list.add(new Employee(2013, "Amit13", "pune", "Maharashtra", "manager", 789, 70952, LocalDate.parse("5-May-2020", formatter)));
        list.add(new Employee(2014, "Sanjay14", "kochi", "Kerala", "manager", 789, 117747, LocalDate.parse("6-Feb-2023", formatter)));
        list.add(new Employee(2015, "Amit15", "indore", "MP", "employee", 456, 76894, LocalDate.parse("23-Apr-2019", formatter)));
        list.add(new Employee(2016, "Sanjay16", "delhi", "Delhi", "employee", 456, 80580, LocalDate.parse("12-Nov-2019", formatter)));
        list.add(new Employee(2017, "Sanjay17", "noida", "UP", "manager", 789, 104211, LocalDate.parse("10-Mar-2024", formatter)));
        list.add(new Employee(2018, "Vikram18", "noida", "UP", "employee", 456, 111303, LocalDate.parse("6-Oct-2019", formatter)));
        list.add(new Employee(2019, "Meera19", "noida", "UP", "employee", 456, 98693, LocalDate.parse("6-Aug-2022", formatter)));
        list.add(new Employee(2020, "Vikram20", "pune", "Maharashtra", "manager", 789, 96727, LocalDate.parse("5-Jun-2017", formatter)));
        list.add(new Employee(2021, "Vikram21", "delhi", "Delhi", "manager", 789, 45794, LocalDate.parse("20-Dec-2021", formatter)));
        list.add(new Employee(2022, "Amit22", "indore", "MP", "employee", 456, 41970, LocalDate.parse("20-Dec-2023", formatter)));
        list.add(new Employee(2023, "Divya23", "pune", "Maharashtra", "manager", 789, 76542, LocalDate.parse("16-Mar-2016", formatter)));
        list.add(new Employee(2024, "Sanjay24", "pune", "Maharashtra", "manager", 789, 110920, LocalDate.parse("27-May-2016", formatter)));
        list.add(new Employee(2025, "Rahul25", "delhi", "Delhi", "manager", 789, 116889, LocalDate.parse("6-Jan-2020", formatter)));
        list.add(new Employee(2026, "Rahul26", "delhi", "Delhi", "manager", 789, 97336, LocalDate.parse("8-Jul-2018", formatter)));
        list.add(new Employee(2027, "Meera27", "delhi", "Delhi", "manager", 789, 109802, LocalDate.parse("21-Jan-2023", formatter)));
        list.add(new Employee(2028, "Sanjay28", "kochi", "Kerala", "manager", 789, 62220, LocalDate.parse("21-Nov-2021", formatter)));
        list.add(new Employee(2029, "Anjali29", "kochi", "Kerala", "employee", 456, 60332, LocalDate.parse("14-Dec-2016", formatter)));
        list.add(new Employee(2030, "Divya30", "delhi", "Delhi", "manager", 789, 119167, LocalDate.parse("17-Jun-2019", formatter)));
        list.add(new Employee(2031, "Neha31", "noida", "UP", "employee", 456, 116853, LocalDate.parse("10-Apr-2016", formatter)));
        list.add(new Employee(2032, "Kavya32", "delhi", "Delhi", "manager", 789, 58556, LocalDate.parse("8-Aug-2016", formatter)));
        list.add(new Employee(2033, "Amit33", "pune", "Maharashtra", "employee", 456, 63008, LocalDate.parse("7-Jan-2018", formatter)));
        list.add(new Employee(2034, "Divya34", "pune", "Maharashtra", "employee", 456, 99342, LocalDate.parse("5-Jan-2022", formatter)));
        list.add(new Employee(2035, "Kavya35", "kochi", "Kerala", "manager", 789, 88149, LocalDate.parse("4-Mar-2019", formatter)));
        list.add(new Employee(2036, "Divya36", "delhi", "Delhi", "employee", 456, 116356, LocalDate.parse("21-Jan-2016", formatter)));
        list.add(new Employee(2037, "Rahul37", "pune", "Maharashtra", "employee", 456, 54841, LocalDate.parse("24-Jul-2023", formatter)));
        list.add(new Employee(2038, "Amit38", "kochi", "Kerala", "employee", 456, 92961, LocalDate.parse("8-Mar-2019", formatter)));
        list.add(new Employee(2039, "Rahul39", "pune", "Maharashtra", "manager", 789, 41969, LocalDate.parse("1-Jul-2016", formatter)));
        list.add(new Employee(2040, "Rahul40", "indore", "MP", "manager", 789, 53844, LocalDate.parse("28-Jun-2016", formatter)));
        list.add(new Employee(2041, "Divya41", "indore", "MP", "manager", 789, 107943, LocalDate.parse("28-May-2017", formatter)));
        list.add(new Employee(2042, "Sanjay42", "indore", "MP", "manager", 789, 45872, LocalDate.parse("12-Feb-2018", formatter)));
        list.add(new Employee(2043, "Neha43", "indore", "MP", "manager", 789, 61430, LocalDate.parse("25-Jan-2021", formatter)));
        list.add(new Employee(2044, "Anjali44", "kochi", "Kerala", "manager", 789, 52564, LocalDate.parse("29-Nov-2020", formatter)));

        return list;
    }

    private static void writeEmployeesToExcel(List<Employee> list, String fileName) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Employees");

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < HEADER.length; i++) headerRow.createCell(i).setCellValue(HEADER[i]);

        int rowIdx = 1;
        for (Employee e : list) {
            Row row = sheet.createRow(rowIdx++);
            row.createCell(0).setCellValue(e.id);
            row.createCell(1).setCellValue(e.name);
            row.createCell(2).setCellValue(e.city);
            row.createCell(3).setCellValue(e.state);
            row.createCell(4).setCellValue(e.category);
            if (e.managerId == null) {
                row.createCell(5).setCellValue("");
            } else {
                row.createCell(5).setCellValue(String.valueOf(e.managerId));
            }
            row.createCell(6).setCellValue(e.salary);
            row.createCell(7).setCellValue(e.doj.format(formatter));
        }

        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            workbook.write(fos);
        }
        workbook.close();
    }

    private static List<Employee> readEmployeesFromExcel(String fileName) throws IOException {
        List<Employee> employees = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(fileName)) {
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rows = sheet.iterator();
            rows.next(); // skip header

            while (rows.hasNext()) {
                Row row = rows.next();
                int id = (int) row.getCell(0).getNumericCellValue();
                String name = row.getCell(1).getStringCellValue();
                String city = row.getCell(2).getStringCellValue();
                String state = row.getCell(3).getStringCellValue();
                String category = row.getCell(4).getStringCellValue();
                String managerStr = row.getCell(5).getStringCellValue();
                Integer managerId = managerStr.isEmpty() ? null : Integer.parseInt(managerStr);
                double salary = row.getCell(6).getNumericCellValue();
                LocalDate doj = LocalDate.parse(row.getCell(7).getStringCellValue(), formatter);
                employees.add(new Employee(id, name, city, state, category, managerId, salary, doj));
            }
            workbook.close();
        }
        return employees;
    }

    private static List<Employee> findGratuityEligibleEmployees(List<Employee> list) {
        return list.stream()
                .filter(e -> ChronoUnit.MONTHS.between(e.doj, LocalDate.now()) > 60)
                .collect(Collectors.toList());
    }

    private static List<Employee> findEmployeesWithHigherSalaryThanManager(List<Employee> list) {
        Map<Integer, Employee> empMap = list.stream().collect(Collectors.toMap(e -> e.id, e -> e));
        return list.stream()
                .filter(e -> e.managerId != null && empMap.containsKey(e.managerId))
                .filter(e -> e.salary > empMap.get(e.managerId).salary)
                .collect(Collectors.toList());
    }

    private static void buildOrgHierarchyJson(List<Employee> list, String fileName) throws IOException {
        Map<Integer, List<Employee>> reports = new HashMap<>();
        Map<Integer, Employee> empMap = new HashMap<>();

        for (Employee e : list) {
            empMap.put(e.id, e);
            if (e.managerId != null) {
                reports.computeIfAbsent(e.managerId, k -> new ArrayList<>()).add(e);
            }
        }

        Employee root = list.stream().filter(e -> e.managerId == null).findFirst().orElse(null);
        ObjectMapper mapper = new ObjectMapper();
        ObjectNode json = buildJsonTree(root, reports, mapper);
        mapper.writerWithDefaultPrettyPrinter().writeValue(new File(fileName), json);
    }

    private static ObjectNode buildJsonTree(Employee e, Map<Integer, List<Employee>> reports, ObjectMapper mapper) {
        ObjectNode node = mapper.createObjectNode();
        node.put("id", e.id);
        node.put("name", e.name);
        node.put("role", capitalize(e.category));
        if (reports.containsKey(e.id)) {
            node.putArray("reportees").addAll(
                    reports.get(e.id).stream()
                            .map(emp -> buildJsonTree(emp, reports, mapper))
                            .collect(Collectors.toList())
            );
        }
        return node;
    }

    private static String capitalize(String s) {
        return s.substring(0, 1).toUpperCase() + s.substring(1);
    }
}
