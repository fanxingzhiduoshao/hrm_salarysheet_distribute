package com.fanxingzhiduoshao.ms.salarysheet.distribute.repository;

import com.fanxingzhiduoshao.ms.salarysheet.distribute.entity.Employee;
import org.springframework.data.repository.CrudRepository;

public interface EmployeeRepository extends CrudRepository<Employee,Integer> {
}
