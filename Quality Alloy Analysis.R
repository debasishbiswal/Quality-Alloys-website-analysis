# Setting the working directory to get the files need for Quality Alloy's Case 

setwd(dir = '/Users/debasishbiswal/Desktop/Web Analytics Case Student Spreadsheet.xls')
getwd()

# Activating libraries needed to import the data and data visualization.

library(readxl)
library(ggplot2)
library(plotly)
library(scales)
library(reshape2)

# Using the readxl package to read excel files
# There are five 5 sheets, we are using those 4 tables that come from the 4 first sheets, and skipping the first 4 rows, 
# to format the dataset as a dataframe. Also, creating 4 tables that come from the demographics information, giving the specific
# range of information.

alloy_wvdf      <- read_excel('Web Analytics Case Student Spreadsheet.xls', sheet = 'Weekly Visits', skip = 4)
alloy_findf     <- read_excel('Web Analytics Case Student Spreadsheet.xls', sheet = 'Financials', skip = 4)
alloy_lbssold   <- read_excel('Web Analytics Case Student Spreadsheet.xls', sheet = 'Lbs. Sold', skip = 4)
alloy_dvdf      <- read_excel('Web Analytics Case Student Spreadsheet.xls', sheet = 'Daily Visits', skip = 4)
alloy_eng       <- read_excel('Web Analytics Case Student Spreadsheet.xls', sheet = 'Demographics', range='B27:C37')
alloy_geo       <- read_excel('Web Analytics Case Student Spreadsheet.xls', sheet = 'Demographics', range='B40:C50')
alloy_brow       <- read_excel('Web Analytics Case Student Spreadsheet.xls', sheet = 'Demographics', range='B54:C64')
alloy_syst       <- read_excel('Web Analytics Case Student Spreadsheet.xls', sheet = 'Demographics', range='B68:C78')

############################# Checking if there is any missing values in the datasets #############################

any(is.na(alloy_dvdf))        #no missing values
any(is.na(alloy_findf))       #no missing values
any(is.na(alloy_lbssold))     #no missing values
any(is.na(alloy_wvdf))        #no missing values
any(is.na(alloy_eng))         #no missing values
any(is.na(alloy_geo))         #no missing values
any(is.na(alloy_brow))        #no missing values
any(is.na(alloy_syst))        #no missing values

# As there is no missing values in any datasets, keep going. 

# Now joining the tables of weekly visits and the financial results. But first, its important to create a column with the periods 
# of the time frame, as will help developping plots and graphs. Also, not only is important to create the periods - initial shakedown, 
# pre promotion, promotion and post promotion -, but also creating an auxiliary variable with value between 1 and 4, to order the data 
# cronologically.

alloy_wvdf$Period_Time_Aux <- c(0)
alloy_wvdf$Period_Time_Aux[1:14] <- 1
alloy_wvdf$Period_Time_Aux[15:35] <- 2
alloy_wvdf$Period_Time_Aux[36:52] <- 3
alloy_wvdf$Period_Time_Aux[53:66] <- 4

################ Changes the columns names ####################

colnames(alloy_brow)  <- c('Browser', 'Visits')
colnames(alloy_eng)   <- c('Searching Engine', 'Visits')
colnames(alloy_geo)   <- c('Region', 'Visits')
colnames(alloy_syst)  <- c('Operating System', 'Visits')


# As the weekly visits and financial sheets have the same quantity of rows and the rows are in order, is possible to use cbind.
# Also, skipping the first column of the Financials table, there is no need 2 columns of week dates.

alloy_bigdf <- cbind(alloy_wvdf, alloy_findf[,2:5])

############################# Creating two new variables, Cost and Profit Margin. #############################

alloy_bigdf$Cost <- alloy_bigdf$Revenue - alloy_bigdf$Profit
alloy_bigdf$Profit_Margin <- alloy_bigdf$Profit/alloy_bigdf$Revenue

############################# Creating plots using ggplots. #############################
############################# Revenue - Linear Regression   #############################

ggplot(alloy_bigdf, aes(alloy_bigdf$`Lbs. Sold`, alloy_bigdf$Revenue)) + geom_point() +
geom_smooth(method = 'lm')+ xlab('Pounds Sold') + ylab('Revenue') +scale_y_continuous(labels = dollar_format())+
scale_x_continuous(labels = scales :: comma)+ theme_classic()  + theme_classic() +
ggtitle('Revenue vs Pound Sold - Linear Regression') +  theme(plot.title = element_text(hjust = 0.5))

############################# Saving the plot above ##################################### 

ggsave('Revenue vs Pound Sold - Linear Regression.png')

################ Checking correlation between revenue and pound sold ####################

cor(alloy_bigdf$`Lbs. Sold`, alloy_bigdf$Revenue)

############################# Revenues Boxplot #############################

ggplot(alloy_bigdf, aes(x = factor(alloy_bigdf$Period_Time_Aux), y = alloy_bigdf$Revenue)) + geom_boxplot() +
  xlab('Period') + ylab('Revenue')+ scale_y_continuous(labels = dollar_format()) + theme_classic() +
  scale_x_discrete(label = c('Initial Shakedown', 'Pre Promotion', 'Promotion', 'Post Promotion'))+
  ggtitle('Revenue Per Period Boxplot') +  theme(plot.title = element_text(hjust = 0.5))

############################# Saving the plot above ##################################### 

ggsave('Revenue Per Period Boxplot.png')

############################# Profit barplot by period   ###############################

ggplot(alloy_bigdf, aes(x = factor(alloy_bigdf$Period_Time_Aux), y = alloy_bigdf$Profit)) + geom_bar(stat = 'identity', fill = 'gray', alpha = 0.8)+
  xlab('Period') + ylab('Profit')+ scale_y_continuous(labels = dollar_format())+ 
  scale_x_discrete(label = c('Initial Shakedown', 'Pre Promotion', 'Promotion', 'Post Promotion')) +
  theme_classic() + ggtitle('Profit per Period') +  theme(plot.title = element_text(hjust = 0.5))

############################# Saving the plot above ##################################### 

ggsave('Profit per Period.png')

##################### Building a table of  inquiries by time ############################

agg_inquiries  <- aggregate(alloy_bigdf$Inquiries, by=list(alloy_bigdf$Period_Time_Aux), FUN=sum)
colnames(agg_inquiries) <- c('Period', 'Total Inquiries')

##################### Building a plot inquiries by time ############################

ggplot(agg_inquiries, aes(x = factor(agg_inquiries$Period), y = agg_inquiries$`Total Inquiries`, group=1)) + geom_line()+
  xlab('Period') + ylab('Total Inquiries')+ geom_line(color = 'red') + geom_point(size = 10) + ylim(60, 140) +
  scale_x_discrete(label = c('Initial Shakedown', 'Pre Promotion', 'Promotion', 'Post Promotion')) +
  theme_classic() + ggtitle('Total Inquiries per Period') +  theme(plot.title = element_text(hjust = 0.5))

############################# Saving the plot above ##################################### 

ggsave('Total Inquiries per Period.png')

################ Checking correlation between inquiries and revenues ####################

cor(alloy_bigdf$Inquiries, alloy_bigdf$Revenue)

################ Plotting pounds solds by week during the weeks analyzed ####################

ggplot(alloy_bigdf, aes(x = alloy_bigdf$`Week (2008-2009)`, y = alloy_bigdf$`Lbs. Sold`, group=1)) + geom_line()+
  xlab('Period') + ylab('Pounds Sold')+ geom_line(color = 'red') + geom_point() + scale_y_continuous(labels = scales :: comma)+
  theme_classic() + ggtitle('Pounds Sold per Week') +  theme(plot.title = element_text(hjust = 0.5)) +
  theme(axis.text.x = element_text(angle = 45, vjust = 1, size = 12, hjust = 1)) + 
  theme(axis.title.x=element_blank(),
      axis.text.x=element_blank(),
      axis.ticks.x=element_blank())

############################# Saving the plot above ##################################### 

ggsave('Pounds Sold per Week.png')

##################### Building a table of  Average Time on Site by time ############################

agg_avgonsite  <- aggregate(alloy_bigdf$`Avg. Time on Site (secs.)`, by=list(alloy_bigdf$Period_Time_Aux), FUN=mean)
colnames(agg_avgonsite) <- c('Period', 'Average Time Spent')

############################# Average Time on Site by Period Boxplot #############################

ggplot(alloy_bigdf, aes(x = factor(alloy_bigdf$Period_Time_Aux), y = alloy_bigdf$`Avg. Time on Site (secs.)`)) + geom_boxplot(fill = 'royal blue') +
  xlab('Period') + ylab('Average Time Spent on Site (secs)')+ theme_classic() +
  scale_x_discrete(label = c('Initial Shakedown', 'Pre Promotion', 'Promotion', 'Post Promotion'))+
  ggtitle('Average Time Spent on Site (secs.) Per Period') +  theme(plot.title = element_text(hjust = 0.5))

############################# Saving the plot above ##################################### 

ggsave('Average Time Spent on Site (secs).png')

############################# Average Time on Site by Period Barplot #############################

ggplot(agg_avgonsite, aes(x = factor(agg_avgonsite$Period), y = agg_avgonsite$`Average Time Spent`, group=1)) + 
  geom_bar(stat = 'identity')+
  xlab('Period') + ylab('Average Time Spent on Site (secs.)')+ 
  scale_x_discrete(label = c('Initial Shakedown', 'Pre Promotion', 'Promotion', 'Post Promotion')) +
  theme_classic() + ggtitle('Average Time Spent on Site (secs.) per Period') +  theme(plot.title = element_text(hjust = 0.5))

############################# Saving the plot above ##################################### 

ggsave('Average Time Spent on Site (secs) Barplot.png')
    
############################# Cost #############################
############################# Revenue - Linear Regression   #############################

ggplot(alloy_bigdf, aes(alloy_bigdf$`Lbs. Sold`, alloy_bigdf$Cost)) + geom_point() +
geom_smooth(method = 'lm')+ xlab('Pounds Sold') + ylab('Cost') +scale_y_continuous(labels = dollar_format()) +
scale_x_continuous(labels = scales :: comma) + theme_classic()+
ggtitle('Cost vs Pound Sold - Linear Regression') +  theme(plot.title = element_text(hjust = 0.5))

############################# Saving the plot above ##################################### 

ggsave('Cost vs Pound Sold - Linear Regression.png')

############################# Building the Correlation ##################################### 

cor(alloy_bigdf$`Lbs. Sold`, alloy_bigdf$Cost)

############################# Bounce Rate vs Visitor - Scatter #############################
############################# Scatterplot Bounce Rate vs Visits #############################

ggplot(alloy_bigdf, aes(alloy_bigdf$`Bounce Rate`, alloy_bigdf$Visits)) + geom_point() + xlab('Bounce Rate') + ylab('Visits')+ theme_classic()

############################# Building Correlation Bounce Rate vs Visits #############################

cor(alloy_bigdf$`Bounce Rate`, alloy_bigdf$Visits)

############################# Bounce Rate vs Visitor - Boxplot #############################

ggplot(alloy_bigdf, aes(x = factor(alloy_bigdf$Period_Time_Aux), y = alloy_bigdf$`Bounce Rate`)) + geom_boxplot() +
  xlab('Period') + ylab('Bounce Rate') + theme_classic() +
  scale_x_discrete(label = c('Initial Shakedown', 'Pre Promotion', 'Promotion', 'Post Promotion')) +
  ggtitle('Bounce Rate per Period') +  theme(plot.title = element_text(hjust = 0.5))

############################# Saving Bounce Rate vs Visitor - Boxplot #############################

ggsave('Bounce Rate per Period.png')

############################# Revenue vs Pound Linear Model #############################

rev_linearmod <- lm(Revenue ~ `Lbs. Sold`, data=alloy_bigdf)  # build linear regression model on full data
print(rev_linearmod)
summary(rev_linearmod)

############################# Revenue vs Page Visit Linear Model #############################

rev_linearmod2 <- lm(Revenue ~ `Pageviews`, data=alloy_bigdf)  # build linear regression model on full data
print(rev_linearmod2)
summary(rev_linearmod2)

############################# Cost vs Pound Linear Model #############################

cost_linearmod <- lm(Cost ~ `Lbs. Sold`, data=alloy_bigdf)  # build linear regression model on full data
print(cost_linearmod)
summary(cost_linearmod)

############################# Profit vs Pound Linear Model #############################

prof_linearmod <- lm(Profit ~ `Lbs. Sold`, data=alloy_bigdf)  # build linear regression model on full data
print(prof_linearmod)
summary(prof_linearmod)

############################# Creating Aggregation Tables #############################

############################# Creating Financials Aggregation Table #############################

agg_profit  <- aggregate(cbind(alloy_bigdf$Revenue, alloy_bigdf$Cost, alloy_bigdf
                               $Profit), by=list(alloy_wvdf$Period_Time, alloy_bigdf$Period_Time_Aux), FUN=sum)
agg_profit$V4 <- sprintf("%1.1f%%",100*agg_profit$V3/ agg_profit$V1)
agg_profit$V1 <- format(agg_profit$V1,scientific=FALSE, big.mark=',')
agg_profit$V2 <- format(agg_profit$V2,scientific=FALSE, big.mark=',')
agg_profit$V3 <- format(agg_profit$V3,scientific=FALSE, big.mark=',')
colnames(agg_profit) <- c('Period', 'Group Number', 'Revenue_Agg','Cost_Agg', 'Profit_Agg', 'Profit Margin')

############################# Creating Pounds per Period Aggregation Table #############################

agg_pounds  <- aggregate(alloy_bigdf$`Lbs. Sold`, by=list(alloy_wvdf$Period_Time, alloy_bigdf$Period_Time_Aux), FUN=mean)
summary(factor(alloy_bigdf$Period_Time_Aux))
summary(str(alloy_bigdf))

############################# Creating Aggregration daily visitis Table #############################

alloy_dvdf$Weekday_num <- c(7,1,2,3,4,5,6)
alloy_dvdf$Day_tipe <- c('Weekend', 'Weekday', 'Weekday', 'Weekday', 'Weekday', 'Weekday', 'Weekend')
agg_visits  <- aggregate(alloy_dvdf$Visits, by=list(alloy_dvdf$Weekday_num), FUN=mean)
colnames(agg_visits) <- c('Weekday', 'Average Visit')

############################# Creating Average Visits per Day Plot #############################

ggplot(agg_visits, aes(x = factor(agg_visits$Weekday), y = agg_visits$`Average Visit`)) + geom_point(size = 7) +
  xlab('Weekday') + ylab('Average Visit')  + ggtitle('Average Visits per Day') +  theme(plot.title = element_text(hjust = 0.5))+
  scale_x_discrete(label = c('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')) +
  theme_classic()

############################# Saving the plot above #############################

ggsave("Average Visit per day.png")

############################# Creating a heatmap #############################
############################# Subsetting to get only the numeric columns, and creating a corrlation matrix #############################

alloy_bigdf_numeric <- alloy_bigdf[,c(2:15)]
corr_mat <- signif(cor(alloy_bigdf_numeric), 2)
col<- colorRampPalette(c("red", "white", "blue"))(20)
heatmap(corr_mat, col=col) 

############################# Adusting the correlation table #############################

corr_mat2 <- melt(corr_mat)

############################# Creating a heatmap #############################

ggplot(corr_mat2, aes(Var2, Var1, fill = value))+
  geom_tile(color = "white")+
  scale_fill_gradient2(low = "red", high = "royal blue", mid = "white", 
                       midpoint = 0, limit = c(-1,1), space = "Lab", 
                       name="Pearson\nCorrelation") +
  theme_minimal()+ # minimal theme
  theme(axis.text.x = element_text(angle = 45, vjust = 1, 
                                   size = 12, hjust = 1))+
  ggtitle('Correlation Matrix') + theme(plot.title = element_text(hjust = 0.5)) +
  coord_fixed()+
  geom_text(aes(Var2, Var1, label = value), color = "black", size = 3) +
  theme(
    axis.title.x = element_blank(),
    axis.title.y = element_blank())

############################# Creating an Pound per Period Aggregation Table #############################

agg_pounds  <- aggregate(alloy_bigdf$`Lbs. Sold`, by=list(alloy_wvdf$Period_Time_Aux, alloy_bigdf$Period_Time_Aux), FUN=mean)

############################# Creating an Average Pound per Period Plot #############################

ggplot(agg_pounds, aes(x = factor(agg_pounds$Group.2), y = agg_pounds$x, group=1)) + 
  geom_bar(stat = 'identity', fill = 'royal blue', alpha = 0.8)+ coord_cartesian(ylim = c(0, 20000)) +
  xlab('Period') + ylab('Pounds Sold') + scale_y_continuous(labels = scales :: comma)+
  scale_x_discrete(label = c('Initial Shakedown', 'Pre Promotion', 'Promotion', 'Post Promotion')) +
  theme_classic() + ggtitle('Average Pounds Sold per Period') +  theme(plot.title = element_text(hjust = 0.5))


############################# APPENDIXES PLOTS #############################

############################# Visits per Searching Engine Plot #############################

alloy_eng$order <- c(1:10)
ggplot(alloy_eng, aes(x = factor(alloy_eng$order), y = alloy_eng$Visits)) + geom_bar(stat = 'identity', fill = 'royal blue', alpha = 0.8)+
  xlab('Period') + ylab('Visitors')+ scale_y_continuous(labels = scales :: comma)+ 
  scale_x_discrete(label = c('google','yahoo','search','msn','aol','ask','live','bing','voila','netscape')) +
  theme_classic() + ggtitle('Web Visitors per Searching Engine') +  theme(plot.title = element_text(hjust = 0.5))

############################# Saving the plot above #############################
ggsave('Web Visitors per Searching Engine.png')

############################# Visits per Browser Plot #############################

alloy_brow$order <- c(1:10)
ggplot(alloy_brow, aes(x = factor(alloy_brow$order), y = alloy_brow$Visits)) + geom_bar(stat = 'identity', fill = 'royal blue', alpha = 0.8)+
  xlab('Period') + ylab('Visitors')+ scale_y_continuous(labels = scales :: comma)+ 
  scale_x_discrete(label = c('Internet Explorer','Firefox','Opera','Safari','Chrome','Mozilla','Netscape','Konqueror','SeaMonkey','Camino')) +
  theme_classic() + ggtitle('Web Visitors per Browser') +  theme(plot.title = element_text(hjust = 0.5))

############################# Saving the plot above #############################
ggsave('Web Visitors per Browser.png')

