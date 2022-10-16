package io.github.yyfcode.fastexcel.demo.entity;

import java.util.Date;

import lombok.Data;

/**
 * Simple JavaBean domain object representing a visit.
 * @author Justice
 */
@Data
public class Visit {

	private Date visitDate;

	private String description;
}
