from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import DataCenter
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer
from officeintegrator.src.com.zoho.officeintegrator.v1 import DocumentDeleteSuccessResponse, \
    InvalidConfigurationException, Authentication, EditPdfParameters
from officeintegrator.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class DeleteDocument:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-pdfeditor-delete-pdf.html
    @staticmethod
    def execute():
        DeleteDocument.init_sdk()
        editDocumentParams = EditPdfParameters()

        # Either use url as document source or attach the document in request body use below methods
        editDocumentParams.set_url('https://demo.office-integrator.com/zdocs/BusinessIntelligence.pdf')

        print('Creating a document to demonstrate document delete api')
        v1Operations = V1Operations()
        response = v1Operations.edit_pdf(editDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CreateDocumentResponse):
                    documentId = str(responseObject.get_document_id())
                    print('Document ID to be deleted : ' + documentId)

                    deleteApiResponse = v1Operations.delete_pdf_document(documentId)

                    if deleteApiResponse is not None:
                        print('Status Code: ' + str(deleteApiResponse.get_status_code()))
                        deleteResponseObject = deleteApiResponse.get_object()

                        if deleteResponseObject is not None:
                            if isinstance(deleteResponseObject, DocumentDeleteSuccessResponse):
                                print('PDF Document delete status : ' + str(deleteResponseObject.get_document_deleted()))
                            elif isinstance(deleteResponseObject, InvalidConfigurationException):
                                print('Invalid configuration exception.')
                                print('Error Code  : ' + str(deleteResponseObject.get_code()))
                                print("Error Message : " + str(deleteResponseObject.get_message()))
                                if deleteResponseObject.get_parameter_name() is not None:
                                    print("Error Parameter Name : " + str(deleteResponseObject.get_parameter_name()))
                                if deleteResponseObject.get_key_name() is not None:
                                    print("Error Key Name : " + str(deleteResponseObject.get_key_name()))
                            else:
                                print('Delete PDF Document Request Failed')
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Document PDF Session Creation Request Failed')

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

DeleteDocument.execute()
