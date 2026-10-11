# Excel-module-end-Assignment
Health Care Data Analysis And Insights
found the number of cells containing "?" in 'medical examination" table and "hospitalisation details" by =COUNTIF(A:K,"?") & =COUNTIF(A:J,"?")
Filled missing values of month by sep using ctrl+H and year by =ROUND(AVERAGE(B:B,0))
Found most frequently occured in smoker by =COUNTIF(J:J,"YES") =COUNTIF(J:J,"NO")  like Hospital tier and city tier by =COUNTIF and replaced blanks with the result by ctrl+H
State id missed filled with "Unknown" by ctrl+H
Split names into 3 columns using TEXT TO COLUMN
Converted "no major surgeries" to "0" by ctrl+H
Inconsistencies in smoker and heart issues was the "capitalisation" changes "yes" to "Yes"
created a new column "weight status" by the condition =IF([@BMI]<18.5,"Under weight",IF([@BMI]<25,"Normal weight",IF([@BMI]<30,"Over weight","Obesity")))
"Diabetic status" by =IF([@HBA1C]<5.7,"Normal",IF([@HBA1C]<6.5,"Prediabetes","Diabetes"))
created a column "date of birth" and used the formula =DATE([@year],MONTH(1&[@month]),[@date])
calculated the age with 8-6-2023 by =DATEDIF([@[Date of Birth]],DATE(2023,6,8),"y")
Changed charges from "general" to "currency" 
created a new sheet as "health care" and copied first column "customer id" from the first table and copied all "headers" after that used "VLOOKUP" to find all the column cells contents
eg:=VLOOKUP(A2,'Customer Names'!A:D,3,FALSE)
created pivot table from insert tab and give field as "cancer history" and "smokers" and from pivot table analysis created pivot chart as pie chart
"any transplants" in raws and "average of HBA1C" and "sum of no:of major surgeries" in values
"diabetic status" in raws "weight status" in column "average of charges" in values and created a column chart 
"state id" in raws "hospital tier" in column "average of charges" in values 
"age" in raws "average of HbA1C" "Average of BMI" in values and created a "line chart" from pivot charts
"AGE" in raws "average of charges" in values
