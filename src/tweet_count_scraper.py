import tweepy
import pandas as pd
import time


bearer_token = "Bearer Token"
queries = pd.read_excel(r"C:\Users\willk\PycharmProjects\DataScienceProject\Config.xlsx")
client = tweepy.Client(bearer_token)

for i in range(len(queries["NFT collection"])):
    year = int(queries["year"][i])
    month = int(queries["month"][i])
    day = 1
    hour = 0
    minutes = 0

    while True:
        if month < 12:
            startTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' \
                        + str(hour) + ':' + str(minutes) + ':00.000Z'
            month += 1
            day = 1
            hour = 0
            endTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' + str(hour) + ':' + str(minutes) + ':00.000Z'
            counts = client.get_all_tweets_count(query=queries["query"][i],
                                                 granularity='day',
                                                 start_time=startTime,
                                                 end_time=endTime)
            df = pd.DataFrame(counts.data)
            file_name = queries["NFT collection"][i]
            file_name = file_name + '_' + str(month-1) + '_' + str(year) + 'Count' + ".csv"
            df.to_csv(file_name, index=False)
            time.sleep(4)
        else:
            startTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' \
                        + str(hour) + ':' + str(minutes) + ':00.000Z'
            year += 1
            month = 1
            day = 1
            hour = 0
            endTime = str(year) + '-' + str(month) + '-' + str(day) + 'T' + str(hour) + ':' + str(minutes) + ':00.000Z'
            counts = client.get_all_tweets_count(query=queries["query"][i],
                                                 granularity='day',
                                                 start_time=startTime,
                                                 end_time=endTime)
            df = pd.DataFrame(counts.data)
            file_name = queries["NFT collection"][i]
            file_name = file_name + '_12_' + str(year-1) + 'Count' + ".csv"
            df.to_csv(file_name, index=False)
            time.sleep(4)
        if month > 2 and year == 2022:
            break
