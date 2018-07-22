package com.fanxingzhiduoshao.ms.salarysheet.distribute.entity;

import lombok.Data;

import javax.persistence.*;

@Entity
@Table
@Data
public class Employee {
    @Id
    @GeneratedValue(strategy=GenerationType.AUTO)
    private int id;
    private String empId;
    private String name;
    private String email;

}
