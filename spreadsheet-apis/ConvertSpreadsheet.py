import os

from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import DataCenter
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.util import StreamWrapper
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer
from officeintegrator.src.com.zoho.officeintegrator.v1 import InvalidConfigurationException, \
    SheetConversionParameters, SheetConversionOutputOptions, FileBodyWrapper, Authentication
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class ConvertSpreadsheet:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/sheet-conversion-api.html
    @staticmethod
    def execute():
        ConvertSpreadsheet.init_sdk()
        sheetConversionParameters = SheetConversionParameters()

        # Either use url as document source or attach the document in request body use below methods
        sheetConversionParameters.set_url('https://demo.office-integrator.com/samples/sheet/Contact_List.xlsx')

        ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        # filePath = ROOT_DIR + "/sample_documents/Contact_List.xlsx"
        # print('Path for source file to be edited : ' + filePath)
        # sheetConversionParameters.set_document(StreamWrapper(file_path=filePath))

        outputOptions = SheetConversionOutputOptions()

        outputOptions.set_format('pdf')
        outputOptions.set_document_name('ConvertedSheet.pdf')

        sheetConversionParameters.set_output_options(outputOptions)

        v1Operations = V1Operations()
        response = v1Operations.convert_sheet(sheetConversionParameters)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, FileBodyWrapper):
                    convertedDocument = responseObject.get_file()

                    if isinstance(convertedDocument, StreamWrapper):
                        outputFileStream = convertedDocument.get_stream()
                        outputFilePath = ROOT_DIR + "/sample_documents/conversion_output.pdf"

                        with open(outputFilePath, 'wb') as outputFileObj:
                            # while True:
                            #     # Read a chunk of data from the input stream
                            #     chunk = outputFileStream.read(1024)  # You can adjust the chunk size as needed
                            #
                            #     # If no more data is read, break the loop
                            #     if not chunk:
                            #         break
                            #
                            #     # Write the chunk of data to the file
                            #     outputFileObj.write(chunk)
                            outputFileObj.write(outputFileStream.content)

                        print("\nCheck converted output file in file path - " + outputFilePath)
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Spreadsheet Conversion Request Failed')

    @staticmethod
    def init_sdk():
        try:
            #Sdk application log configuration
            logger = Logger.get_instance(Logger.Levels.INFO, "./logs.txt")
            #Update this apikey with your own apikey signed up in office integrator service
            auth = Auth.Builder().add_param("apikey", "2ae438cf864488657cc9754a27daa480").authentication_schema(Authentication.TokenFlow()).build()
            tokens = [ auth ]
            # Refer this help page for api end point domain details -  https://www.zoho.com/officeintegrator/api/v1/getting-started.html
            environment = DataCenter.Production("https://api.office-integrator.com")

            Initializer.initialize(environment, tokens,None, None, logger, None)
        except SDKException as ex:
            print(ex.code)

ConvertSpreadsheet.execute()