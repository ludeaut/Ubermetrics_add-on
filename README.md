# Ubermetrics_add-on
Project done in Summer 2019 while I was part of INTK as a software development intern. <br/>
Website here: https://intk.com/en <br/>
This internship was part of my three-year Master's Degree in Engineering that I follow at Grenoble INP - Ensimag.

Ubermetrics is an add-on which can be used on Google Sheets to fetch Google Analytics and Google Ads data.
From the user point of view, the add-on contains two functions:
* UberAnalytics fetching Google Analytics data. <br/>
  The parameters required are a view ID, a data range and a metric but it is also possible to:
    * Specify multiple metrics instead of only one,
    * Add dimensions,
    * Add filters, multiple filters can be added using commas and semicolons to represent ‘or’ and ‘and’ operators,
    * Specify a sort order with a sort field and sort direction,
    * Set a maximum to the number of results returned.
* UberAds fetching Google Ads data. <br/>
  The parameters required are a manager account ID, an account ID, a data range, a field, a resource and a segment of type     DATE as last condition but it is also possible to:
    * Specify multiple fields instead of only one,
    * Add conditions,
    * Specify a sort order with a sort field and sort direction,
    * Set a maximum to the number of results returned.

A help documentation for UberAnalytics and UberAds is displayed in Google Sheets. <br/>
The user has also the possibility to create a setup sheet from a custom menu. The new sheet called Setup UberMetrics contains a link to the add-on documentation (link only available to the INTK's employees and can be found below) and the possible fields to be filled for UberAnalytics and UberAds at the exception of the client ID (respectively view ID and the account ID).


UberAnalytics
-----------------

Dimensions & Metrics: https://ga-dev-tools.appspot.com/dimensions-metrics-explorer/ <br/>
Possible dates: https://developers.google.com/analytics/devguides/reporting/core/v3/reference#startDate

How to have data from last month
* startDate: =EOMONTH(TODAY(), -2) + 1
* endDate: =EOMONTH(TODAY(), -1)

How to have data from last week (Monday to Sunday)
* startDate: =TODAY() - WEEKDAY(TODAY(),2) - 6
* endDate: =TODAY() - WEEKDAY(TODAY(),2)

How to have data from last week (Monday to Friday)
* startDate: =TODAY() - WEEKDAY(TODAY(),2) - 6
* endDate: =TODAY() - WEEKDAY(TODAY(),2) - 2

UberAds
-------------

Manager Account ID and Account ID need to be given as a string with this format XXX-XXX-XXXX or with this one XXXXXXXXXX. <br/>
Always add "segments.date" in last to the conditions, it deals with the date range. <br/>
The only exception is if you specify a date segment (segments.week e.g.) in the fields, 
then add the date segment used in last to the conditions. <br/>
You can specify metrics and segments in the metrics field. <br/>
Metrics: https://developers.google.com/google-ads/api/fields/v2/metrics. <br/>
Segments: https://developers.google.com/google-ads/api/fields/v2/segments. <br/>
To find the resources you can select them with, look in the 'Selectable With' field of a table. <br/>
Possible date ranges: https://developers.google.com/google-ads/api/docs/query/date-ranges.

Note <br/>
It's possible to see some discrepancies between the results returned by UberAds and the results displayed in the Google Ads
interface if the concerned account has Smart campaigns.
