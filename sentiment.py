from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
import tweepy
import xlsxwriter
# xlsxwriter is used for writing into Excel workbook
import twitter_credentials
# Consumer key, consumer secret, access key and access secret for accessing Twitter API are stored in this file

auth = tweepy.OAuthHandler(twitter_credentials.consumer_key, twitter_credentials.consumer_secret)
auth.set_access_token(twitter_credentials.access_key, twitter_credentials.access_secret)
api = tweepy.API(auth)

#we are trying to get 150 tweets and store them into an excel file in which a sheet is named Tweets

tweet_limit=150
wb=xlsxwriter.Workbook('tweetsentiment.xlsx')
# We create a workbook named tweetsentiment.xlsx to store the tweets and the corresponding sentiment scores
ws=wb.add_worksheet('Tweets')

# an empty array in whhich the tweets will load
Tweet=[]
Tweet = tweepy.Cursor(api.search, q='ISRO', lang='en').items(tweet_limit)
#print(adobeTweet.text)

# sentiment is a variable that stores the average sentiment of all the tweets

sentiment=0
analyzer=SentimentIntensityAnalyzer()

row=0
col=0
ws.write(row,col,'Tweets')
ws.write(row,col+1,'Positive Sentiment')
ws.write(row,col+2,'Negative Sentiment')
ws.write(row,col+3,'Neutral')
ws.write(row,col+4,'Compound Score')
row+=1

for tweet in Tweet:    # We run a for loop
   vs=analyzer.polarity_scores(tweet.text)
   # The variable vs stores the sentiment of the tweet under consideration at the time
   #print(vs)
   sentiment=sentiment+vs['compound']
   # Writing into Excel file tweetsentiment.xlsx
   ws.write_string(row,col,tweet.text)
   ws.write(row, col + 1, vs['pos'])
   ws.write(row, col + 2, vs['neg'])
   ws.write(row, col + 3, vs['neu'])
   ws.write(row, col + 4, vs['compound'])
   row+=1

wb.close()
sentiment=sentiment/tweet_limit
print("Sentiment score ------>   ",sentiment)
