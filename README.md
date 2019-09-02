# Ubermetrics_add-on
Project done in Summer 2019 while I was part of INTK as a software development intern.
Website here: https://intk.com/en
This internship was part of my three-year Master's Degree in Engineering that I follow at Grenoble INP - Ensimag.


UberAnalytics
-----------------

Dimensions & Metrics: https://ga-dev-tools.appspot.com/dimensions-metrics-explorer/
Possible dates: https://developers.google.com/analytics/devguides/reporting/core/v3/reference#startDate

How to have data from last month
startDate: =EOMONTH(TODAY(), -2) + 1
endDate: =EOMONTH(TODAY(), -1)

How to have data from last week (Monday to Sunday)
startDate: =TODAY() - WEEKDAY(TODAY(),2) - 6
endDate: =TODAY() - WEEKDAY(TODAY(),2)

How to have data from last week (Monday to Friday)
startDate: =TODAY() - WEEKDAY(TODAY(),2) - 6
endDate: =TODAY() - WEEKDAY(TODAY(),2) - 2

UberAds
-------------

Manager Account ID and Account ID need to be given as a string with this format XXX-XXX-XXXX or with this one XXXXXXXXXX.
Always add "segments.date" in last to the conditions, it deals with the date range.
The only exception is if you specify a date segment (segments.week e.g.) in the fields,
then add the date segment used in last to the conditions.
You can specify metrics and segments in the metrics field.
Metrics: https://developers.google.com/google-ads/api/fields/v2/metrics.
Segments: https://developers.google.com/google-ads/api/fields/v2/segments.
To find the resources you can select them with, look in the 'Selectable With' field of a table.
Possible date ranges: https://developers.google.com/google-ads/api/docs/query/date-ranges.

Note
It's possible to see some discrepancies between the results returned by UberAds and the results displayed in the Google Ads
interface if the concerned account has Smart campaigns.
