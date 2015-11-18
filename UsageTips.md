# Introduction #

You should feel free to use this workbook in the way that works best for you, but here are a few technique that I think work well.

# Setting up the budget #

The weekly budgets are calculated based on the monthly budgets you enter at the top of the Budget worksheet. If you budget some or all of your tags on a weekly basis then you could enter your weekly budget amounts into the weekly budgets row rather than relying on the formulas that come with the workbook.

You'll notice that there are two total columns, the first is an all-inclusive total, and the second is a total of all but one of the tag columns. I have found that I don't have direct control over many of the things I tag as bills so the second total column lets me see how much I'm over or under on things that I can actually control during the month.

If you need more tags, you can copy the existing formulas into more columns. For more detailed instructions see this [post on Wesabe Groups](https://www.wesabe.com/groups/50-wesabe-api-developers/discussions/2929-excel-add-in#comment_25972).

# Understanding the spending analysis #

Each of the Weekly and Monthly worksheets displays information about how your spending relates to your budget on the date in the first column. What works best for me is to use Monthly Difference worksheet the most. I look at today's numbers to see if I'm over or under budget, and I look at the numbers for the last day of the month to see how much I have left for each tag. When I'm over budget for a tag, I also look at the Monthly Spending Percent worksheet to get a sense of how far over budget I am in proportion to the budget for that category.

To help you get your bearings, the row for today's date will display as bold and highlighted in light yellow. Amounts that are over budget are colored red to draw your attention.

# Changing the dates #

The workbook is setup to show you the three most recent months, ordered from the most recent day to the least recent day. If you want to view dates other than these then change the dates on the Monthly Spending worksheet; all of the other worksheets copy their dates from the values on the Monthly Spending worksheet. You can delete rows if you would like to see less than three months at a time. If you would like to see more than three months at a time, use the Fill command to extend the formulas in each of the columns on each of the worksheets.

For the Weekly worksheets, weeks start on Monday and are always made up of seven days. As a result, the first and last week of the year may overlap into the preceding or following year. This week numbering system is part of the [ISO 8601](http://en.wikipedia.org/wiki/ISO_8601#Week_dates) standard. Let me know if you would be interested in a different way of splitting up weeks and I will consider making the workbook more flexible.

# Removing your data #

If you make an improvement to the workbook you may want to share it with a friend, contribute it to this project, or email it to yourself for use on another computer. In all of those scenarios, you probably want to clear your transaction data out of the workbook before sending it. After clearing your data and sending the workbook, you can get your transaction data into the workbook again by Ctrl+Shift+D.

To clear your Wesabe data from the workbook, go to the Transactions worksheet and select all of the data on the worksheet. Press the Delete key to clear the cell contents. You should receive a prompt that says, "The range you deleted is associated with a query that retrieves data from an external source. Do you want to delete the query in addition to the range?" Click the No button.