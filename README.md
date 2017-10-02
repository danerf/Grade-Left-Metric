# Grade-Left-Metric

Inspired by "Improving Grading Consistency through Grade Lift Reporting" by Ido Millet, we define the Grade Lift Metric as the difference between average class grade and average core curriculum GPA of the class. This metric gives us an assessment of how harsh or lenient the grading was for the given course as compared to the average core curriculum GPA, which is the GPA in all of the courses within each spreasheet. That is, the average GPA for all courses in a given spreadsheet is the core curriculum GPA for that cohort.

Here, an implementation in VBA, namely the Grade_Lift_Metric(), extracts data from all .xlsx files in the directory specified by the variable "Path" by looping through the Excel workbooks and all of their spreadsheets. Then, the Macro Grade_Lift_Metric() computes core curriculum GPAs and subsequently Grade Lift Metricts for each course storing results in arrays.

The last thing the Macro does is generate a chart comparing Grade Lift Metrics of different courses within a department and a cohort to assess grading consistency. 

If the Grade Lift Metric for a given course is positive, then the grading for that class was lenient compared to other classes within the department. If it is negative then it was more harshly graded.

An example of the result of running the code Grade_Lift_Metric() is shown in the word document called GradeLiftCharts.docx

Grades.xlsm is a template for storing student grades, it is where Grade_Lift_Metric() obtains extracts data.
