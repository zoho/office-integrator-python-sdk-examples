from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import DataCenter
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer
from officeintegrator.src.com.zoho.officeintegrator.v1 import DocumentMeta, InvalidConfigurationException, \
    Authentication, EditPdfParameters
from officeintegrator.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class GetPdfDocumentDetails:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-pdfeditor-document-info.html
    @staticmethod
    def execute():
        GetPdfDocumentDetails.init_sdk()
        editDocumentParams = EditPdfParameters()

        # Either use url as document source or attach the document in request body use below methods
        editDocumentParams.set_url('https://demo.office-integrator.com/zdocs/BusinessIntelligence.pdf')

        print('Creating a pdf document session to demonstrate get all document information api')
        v1Operations = V1Operations()
        response = v1Operations.edit_pdf(editDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CreateDocumentResponse):
                    documentId = str(responseObject.get_document_id())
                    print('Created Document ID : ' + documentId)

                    documentInfoResponse = v1Operations.get_document_info(documentId)

                    if documentInfoResponse is not None:
                        print('Status Code: ' + str(documentInfoResponse.get_status_code()))
                        documentInfoObj = documentInfoResponse.get_object()

                        if documentInfoObj is not None:
                            if isinstance(documentInfoObj, DocumentMeta):
                                print('Pdf Document ID : ' + str(documentInfoObj.get_document_id()))
                                print('Pdf Document Name : ' + str(documentInfoObj.get_document_name()))
                                print('Pdf Document Type : ' + str(documentInfoObj.get_document_type()))
                                print('Pdf Document Collaborators Count : ' + str(documentInfoObj.get_collaborators_count()))
                                print('Pdf Document Created Time : ' + str(documentInfoObj.get_created_time()))
                                print('Pdf Document Created Timestamp : ' + str(documentInfoObj.get_created_time_ms()))
                                print('Pdf Document Expiry Time : ' + str(documentInfoObj.get_expires_on()))
                                print('Pdf Document Expiry Timestamp : ' + str(documentInfoObj.get_expires_on_ms()))
                            elif isinstance(documentInfoObj, InvalidConfigurationException):
                                print('Invalid configuration exception.')
                                print('Error Code  : ' + str(documentInfoObj.get_code()))
                                print("Error Message : " + str(documentInfoObj.get_message()))
                                if documentInfoObj.get_parameter_name() is not None:
                                    print("Error Parameter Name : " + str(documentInfoObj.get_parameter_name()))
                                if documentInfoObj.get_key_name() is not None:
                                    print("Error Key Name : " + str(documentInfoObj.get_key_name()))
                            else:
                                print('Get Document Details API Request Failed')
            elif isinstance(responseObject, InvalidConfigurationException):
                print('Invalid configuration exception.')
                print('Error Code  : ' + str(responseObject.get_code()))
                print("Error Message : " + str(responseObject.get_message()))
                if responseObject.get_parameter_name() is not None:
                    print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                if responseObject.get_key_name() is not None:
                    print("Error Key Name : " + str(responseObject.get_key_name()))
            else:
                print('Create Pdf Document Session Request Failed')

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

GetPdfDocumentDetails.execute()