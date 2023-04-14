import pandas as pd

def swap_rows(start, end):

    # FUNCTION TO ARRANGE INDIVIDUALS WITH SAME TOTAL SUM OF STATEMENTS AND REASONS

    name_list = []
    for i in range(start, end+1):
        val = df2[df2['Rank'] == i]['name'].values[0].lower().replace("_","");
        name_list.append(val)
    name_list.sort()
    df_temp = df2.copy()
    for i in range(0,len(name_list)):
        curr = name_list[i]
        put_in_rank = 0
        for j in range(start,end+1):
            temp = df2[df2['Rank'] == j]['name'].values[0].lower().replace("_","");
            if(temp == curr):
                put_in_rank = j
                break
        row_index_1 = df2[df2['Rank'] == start+i].index[0]
        row_index_2 = df2[df2['Rank'] == put_in_rank].index[0]
        df_temp.loc[row_index_1, :] = df2.loc[row_index_2, :]
    df2.update(df_temp)


# ********************************************CODE TO PRODUCE INDIVIDUAL LEADERBORD *******************************************


# reading input excel sheets into df1 and df2
df1 = pd.read_excel('input1.xlsx')
df2 = pd.read_excel('input2.xlsx')
start = 0
end = 1
l = []

# adding a field names total_score and sorting on the basis of total_score
df2['total_score'] = df2['total_statements'] + df2['total_reasons']
df2 = df2.sort_values(by=["total_score"], ascending=False)
df2.insert(0, 'Rank', range(1, 1 + len(df2)))

# sorting alphabetically if the value of total_score is same
for i in range(1,len(df1)):
    val = df2[df2['Rank'] == i]['total_score'].values[0];
    l.append(val)

# checking which ranges have the same value for total_score
while start<len(df1):
    if end >= len(df1)-1:
        break
    if l[end] == l[start]:
        end = end+1
    else:
        if end-start > 1:
            swap_rows(start+1, end)
        start = end
if end-start > 1:
    swap_rows(start+1, end)

# droping unnessasry fields and modifying field names, center alignments and saving the excel into file name output1.xlsx
df2 = df2.drop("S No", axis='columns')
df2 = df2.drop("total_score", axis='columns')
df2.rename(columns = {'name':'Name'}, inplace = True)
df2.rename(columns = {'uid':'UID'}, inplace = True)
df2.rename(columns = {'total_statements':'No. of Statements'}, inplace = True)
df2.rename(columns = {'total_reasons':'No. of Reasons'}, inplace = True)
print(df2) # this is the individual ranking dataframe

# SAVING THE RESULT IN THE INDIVIDUAL LEADERBOARD EXCEL FILE
file_name = 'individual_leaderboard.xlsx'
def align_center(x):
    return ['text-align: center' for x in x]
with pd.ExcelWriter(file_name) as writer:
    df2.style.apply(align_center, axis=0).to_excel(
        writer,
        index=False
    )





# ********************************************CODE TO PRODUCE TEAM WISE LEADERBORD *******************************************


# merging and grouping dataframes

merged_df = pd.merge(df1, df2, on = 'Name')
grouped = merged_df.groupby('Team Name')
mean1 = grouped['No. of Statements'].mean().round(2)
mean2 = grouped['No. of Reasons'].mean().round(2)
team_wise_leaderboard = pd.concat([mean1, mean2], axis=1)
team_wise_leaderboard.columns = ['Average Statements', 'Average Reasons']

# Add the team name to the DataFrame and finding average values of statements and reasons
team_wise_leaderboard['Thinking Teams Leaderboard'] = team_wise_leaderboard.index
column_name = 'Thinking Teams Leaderboard'
team_wise_leaderboard = pd.concat([team_wise_leaderboard[column_name], team_wise_leaderboard.drop(column_name, axis=1)], axis=1)


# sorting the dataframe on the basis of sum of average statements and average reasons and adding ranking coloumn
team_wise_leaderboard['total_score'] = team_wise_leaderboard['Average Statements'] + team_wise_leaderboard['Average Reasons']
team_wise_leaderboard = team_wise_leaderboard.sort_values(by=["total_score"], ascending=False)
team_wise_leaderboard.insert(0, 'Rank', range(1, 1 + len(team_wise_leaderboard)))

team_wise_leaderboard = team_wise_leaderboard.drop("total_score", axis='columns')       # dropping total_score coloumn

print(team_wise_leaderboard) # this is Team Wise Leaderboard dataframe


# SAVING THE RESULT IN THE TEAMWISE LEADERBOARD EXCEL FILE
file_name = 'teamwise_leaderboard.xlsx'
def align_center(x):
    return ['text-align: center' for x in x]
with pd.ExcelWriter(file_name) as writer:
    team_wise_leaderboard.style.apply(align_center, axis=0).to_excel(
        writer,
        index=False
    )

