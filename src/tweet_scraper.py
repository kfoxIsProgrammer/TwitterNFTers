import tweepy
import pandas as pd
import time


bearer_token = "Bearer Token"
queries = pd.read_excel(r"C:\Users\willk\PycharmProjects\DataScienceProject\Config copy.xlsx")
client = tweepy.Client(bearer_token)
days_in_month = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

for i in range(len(queries["NFT collection"])):
    year = int(queries["year"][i])
    month = int(queries["month"][i])
    day = 1
    hour = 0
    minutes = 0
    newDF = True

    while True:
        if hour < 23:
            startTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' \
                        + str(hour) + ':' + str(minutes) + ':00.000Z'
            endTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' \
                      + str(hour+1) + ':' + str(minutes) + ':00.000Z'
            hour += 1
            counts = client.search_all_tweets(query=queries["query"][i],
                                              start_time=startTime,
                                              max_results=500,
                                              end_time=endTime,
                                              tweet_fields="text,attachments,public_metrics,created_at")
            if newDF:
                df = pd.DataFrame(counts.data)
                newDF = False
            else:
                df1 = pd.DataFrame(counts.data)
                df = pd.concat([df, df1])
            time.sleep(3)

        elif day < days_in_month[month-1]:
            startTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' \
                        + str(hour) + ':' + str(minutes) + ':00.000Z'
            day += 1
            hour = 0
            endTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' \
                      + str(hour) + ':' + str(minutes) + ':00.000Z'
            counts = client.search_all_tweets(query=queries["query"][i],
                                              start_time=startTime,
                                              max_results=500,
                                              end_time=endTime,
                                              tweet_fields="text,attachments,public_metrics,created_at")
            if newDF:
                df = pd.DataFrame(counts.data)
                newDF = False
            else:
                df1 = pd.DataFrame(counts.data)
                df = pd.concat([df, df1])
            time.sleep(3)

        elif month < 12:
            startTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' \
                        + str(hour) + ':' + str(minutes) + ':00.000Z'
            month += 1
            day = 1
            hour = 0
            endTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' + str(hour) + ':' + str(minutes) + ':00.000Z'
            counts = client.search_all_tweets(query=queries["query"][i],
                                              start_time=startTime,
                                              max_results=500,
                                              end_time=endTime,
                                              tweet_fields="text,attachments,public_metrics,created_at")
            df1 = pd.DataFrame(counts.data)
            df = pd.concat([df, df1])
            file_name = queries["NFT collection"][i]
            file_name = file_name + '_' + str(month-1) + '_' + str(year) + 'Tweets' + ".csv"
            df.to_csv(file_name, index=False)
            newDF = True
            time.sleep(5)
        else:
            startTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' \
                        + str(hour) + ':' + str(minutes) + ':00.000Z'
            year += 1
            month = 1
            day = 1
            hour = 0
            endTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' + str(hour) + ':' + str(minutes) + ':00.000Z'
            counts = client.search_all_tweets(query=queries["query"][i],
                                              start_time=startTime,
                                              max_results=500,
                                              end_time=endTime,
                                              tweet_fields="text,attachments,public_metrics,created_at")
            df1 = pd.DataFrame(counts.data)
            df = pd.concat([df, df1])

            file_name = queries["NFT collection"][i]
            file_name = file_name + '_12' + '_' + str(year-1) + 'Tweets' + ".csv"
            df.to_csv(file_name, index=False)
            time.sleep(3)

        if year == 2022 and month ==3:
            break
