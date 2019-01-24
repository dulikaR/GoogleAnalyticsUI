
from apiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from datetime import timedelta
from io import StringIO
import csv
###########
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
###########



class Analytics():

    SCOPES = ['https://www.googleapis.com/auth/analytics.readonly']
    KEY_FILE_LOCATION = 'client_secrets.json'
    VIEW_ID = '161070304'

    #credentials & authenticating
    def initialize_analyticsreporting(self):

      credentials = ServiceAccountCredentials.from_json_keyfile_name(
          self.KEY_FILE_LOCATION, self.SCOPES)

      # Build the service object.
      analytics = build('analyticsreporting', 'v4', credentials=credentials)

      return analytics


    #send inputs (metrics & dimentions) to api
    def get_report(self,analytics,dimentions,metrics):

    ###############metrics######################

      name_metric = metrics

      ga_metric = []
      for x in range(len(name_metric)):
          ga_metric.append('expression')

      metric = [{q: z} for q, z in zip(ga_metric, name_metric)]


    ####################dimentions#####################
      name_dim = dimentions

      ga_dim = []
      for x in range(len(name_dim)):
          ga_dim.append('name')

      dimen = [{k: v} for k, v in zip(ga_dim, name_dim)]

    #######################################################


      return analytics.reports().batchGet(
          body={
            'reportRequests': [
            {
              'viewId': self.VIEW_ID,
              'dateRanges': [{'startDate': datetime.strftime(datetime.now() - timedelta(days = 30),'%Y-%m-%d'), 'endDate': datetime.strftime(datetime.now(),'%Y-%m-%d')}],
              # 'metrics': [{'expression': 'ga:organicSearches'}],
              'metrics': metric,
                # 'dimensions': [{'name': 'ga:referralPath'},{'name': 'ga:fullReferrer'},{'name': 'ga:campaign'},{'name': 'ga:source'},{'name': 'ga:medium'},{'name': 'ga:sourceMedium'},{'name': 'ga:keyword'}]
              'dimensions': dimen
            }]
          }
      ).execute()


    def send_mail(self,exporting_file,person,product):

        person_trim = person.split('@')

        TOADDR = [str(person)]
        CCADDR = ['malakag@masholdings.com', 'dulikar@masholdings.com']

        isTls = True
        msg = MIMEMultipart()
        msg['From'] = "pixeldatabot@gmail.com"
        msg['To'] = ', '.join(TOADDR)
        msg['Cc'] = ', '.join(CCADDR)
        msg['Date'] = formatdate(localtime = True)
        msg['Subject'] = "Google Analytics"
        msg.attach(MIMEText("HI "+str(person_trim[0])+", \n \n  Attached file contains result set for uploaded  list from "+product+" \n \n Cipher \n MAS PIXEL"))


        part = MIMEBase('application', "octet-stream")
        part.set_payload(exporting_file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="GoogleAnalytics.csv"')
        msg.attach(part)

        #context = ssl.SSLContext(ssl.PROTOCOL_SSLv3)
        #SSL connection only working on Python 3+
        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        if isTls:
            smtp.starttls()
        smtp.login("pixeldatabot@gmail.com", "welcome@123")
        smtp.sendmail("pixeldatabot@gmail.com",  TOADDR + CCADDR, msg.as_string())
        smtp.quit()


    #creating excel file from filtered results
    def excel_file(self,final_list):

        #Transpose list of lists
        final_list = list(map(list, zip(*final_list)))

        google_trans_io = StringIO()
        cw_trans = csv.writer(google_trans_io)
        cw_trans.writerows(final_list)
        google_trans_io.seek(0)

        return google_trans_io


    #filter the response from api
    def filter(self,metric_input,dim_input,response):
        # metric_input = [users, transactions]
        # dim_input = [referralPath, fullReferrer, campaign, source, medium, sourceMedium, keyword]

        column_length = len(metric_input) + len(dim_input)

        columns = [[] for i in range(column_length)]

        for report in response.get('reports', []):
            columnHeader = report.get('columnHeader', {})
            dimensionHeaders = columnHeader.get('dimensions', [])
            metricHeaders = columnHeader.get('metricHeader', {}).get('metricHeaderEntries', [])
            metricHeaders_list = []

            for one_metric_head in metricHeaders:
                metricHeaders_list.append(one_metric_head.get('name', []))

            #adding headers to the excel list
            column_headers = metricHeaders_list + dimensionHeaders
            count_headers = 0
            for column_headers_one in column_headers:
                columns[count_headers].append(column_headers_one)
                count_headers+=1

            #adding rows to the excel list
            for row in report.get('data', {}).get('rows', []):
                dimensions = row.get('dimensions', [])
                dateRangeValues = row.get('metrics', [])
                metrics = dateRangeValues[0].get('values', [])
                all_column_values =  metrics + dimensions

                #adding row by row
                count = 0
                for one_column in all_column_values:
                    columns[count].append(one_column)
                    count += 1

                print("finishOne")
            print("finishTwo")
        print("finishThree")
        return columns


    def main(self,email,startDate,endDate,dimentions,metrics):
      analytics = self.initialize_analyticsreporting()
      response = self.get_report(analytics,dimentions,metrics)
      columns_list = self.filter(dimentions,metrics,response)
      final_excel = self.excel_file(columns_list)
      self.send_mail(final_excel, email, "Google Analytics")


