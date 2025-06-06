from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import DataCenter
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer
from officeintegrator.src.com.zoho.officeintegrator.v1 import SessionInfo, SessionMeta, SessionUserInfo, \
    InvalidConfigurationException, Authentication, EditPdfParameters
from officeintegrator.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class GetPdfSessionDetails:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-pdfeditor-session-details.html
    @staticmethod
    def execute():
        GetPdfSessionDetails.init_sdk()
        editDocumentParams = EditPdfParameters()

        # Either use url as document source or attach the document in request body use below methods
        editDocumentParams.set_url('https://demo.office-integrator.com/zdocs/BusinessIntelligence.pdf')

        v1Operations = V1Operations()
        response = v1Operations.edit_pdf(editDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CreateDocumentResponse):
                    sessionId = str(responseObject.get_session_id())
                    print('Created Document Session ID : ' + sessionId)

                    sessionInfoResponse = v1Operations.get_session(sessionId)

                    if sessionInfoResponse is not None:
                        print('Status Code: ' + str(sessionInfoResponse.get_status_code()))
                        sessionInfoObj = sessionInfoResponse.get_object()

                        if sessionInfoObj is not None:
                            if isinstance(sessionInfoObj, SessionMeta):
                                print('Session Status : ' + str(sessionInfoObj.get_status()))

                                sessionInfo = sessionInfoObj.get_info()

                                if isinstance(sessionInfo, SessionInfo):
                                    print('Session User ID : ' + str(sessionInfo.get_session_url()))

                                sessionUserInfo = sessionInfoObj.get_user_info()

                                if isinstance(sessionUserInfo, SessionUserInfo):
                                    print('Session User ID : ' + str(sessionUserInfo.get_user_id()))
                                    print('Session Display Name : ' + str(sessionUserInfo.get_display_name()))
                            elif isinstance(sessionInfoObj, InvalidConfigurationException):
                                print('Invalid configuration exception.')
                                print('Error Code  : ' + str(sessionInfoObj.get_code()))
                                print("Error Message : " + str(sessionInfoObj.get_message()))
                                if sessionInfoObj.get_parameter_name() is not None:
                                    print("Error Parameter Name : " + str(sessionInfoObj.get_parameter_name()))
                                if sessionInfoObj.get_key_name() is not None:
                                    print("Error Key Name : " + str(sessionInfoObj.get_key_name()))
                            else:
                                print('Get Session Details Request Failed')
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Create PDF Document Session Request Failed')

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

GetPdfSessionDetails.execute()