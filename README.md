# Menstrual Symptoms Analysis
I'm just trying to find whether the birth controls and menstrual hormones have been f**king me up and how have they been f-king me up

<u>Tools:</u> python, jupyter notebook, git, github 

<u>Libraries:</u> see requirements.txt

**Top 3 Most Appeared Symptoms from 2022-2026:**

<img src="plots/top_3_symptoms_all_time.png" alt="top_3_symptoms" width="400">

**Cycle and Period Length**

The plot under compares my average period and cycle length with the regular period and cycle length according to [Mayo Clinic](https://www.mayoclinic.org/healthy-lifestyle/womens-health/in-depth/menstrual-cycle/art-20047186). Note that I was on Kyleena for this whole time, and I’ve noticed my menstrual cycle has been prolonged since starting Kyleena. 

Orange see-through area is the window for regular cycle length, my cycle length is the orange line with dots. So you see that my cycle length (orange stuffs in the plot) is off the f**king charts thanks to the side effects of birth control. 

<img src="plots/cycle_analysis.png" alt="cycle_analysis" width="500">

**Period Intensity**

The plot under shows the period intensity for each day for all my periods. Distinct values in data are 1-4 where 1 represents light, 2 for medium, 3 for heavy and 4 for blood clots. My period intensity has been pretty light overall, and I've been lazy logging my period intensity for a good while. And also can I say this is the kind of overview I don't get from Fl* while I'm paying 169NOK for my yearly subscription. #whatamipayingfor

<img src="plots/period_intensity_heatmap.png" alt="period_intensity_heatmap" width="700">


**Period Symptoms Correlations**

I chose top 12 most common symptoms, this plot tells me stuffs like if I'm nausous and I will be angry as well, or if I'm angrying then I will be panicing as well(??) Tbh, what is this hormonal hell here? 

<img src="plots/symptom_correlations.png" alt="symptoms_correlations" width="600">

**Top 3 Most Frequent Symptoms per Cycle**
Correlation Map of the Three Most Common Symptoms Across All Cycles. Red and 1 means this symptoms is one of the top 3 in this cycle. Now you see I'm cramping (DrawingPain) almost all the time when I'm on period. 
<img src="plots/top_3_symptoms_heatmap.png" alt="top_3_symptoms_heatmap" width="900">
