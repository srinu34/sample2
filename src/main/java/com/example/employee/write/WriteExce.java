package com.example.employee.write;


import java.util.Map;
import java.util.Scanner;
import java.util.TreeMap;
import com.example.entity.Employee;

public class WriteExce {

	public static void main(String[] args) 
	{
		
		Scanner sc = new Scanner(System.in);
		Map<Integer, Object> data = new TreeMap<>();
		System.out.println("Enter Number of Employees:");
		int n = sc.nextInt();
		for (int i = 0; i < n; i++) {
			System.out.println("Enter the employee id: ");
			int id = sc.nextInt();
			System.out.println("employee name:");
			String name = sc.next();
			System.out.println("Location");
			String location = sc.next();
			Employee emp = new Employee(id, name, location);
			data.put(i, emp);
		}
		CreateXl.createXl(data);
	
	}
}
