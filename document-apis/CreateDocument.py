import time

from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import DataCenter
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer
from officeintegrator.src.com.zoho.officeintegrator.v1 import DocumentInfo, UserInfo, Margin, DocumentDefaults, \
    EditorSettings, UiOptions, CallbackSettings, InvalidConfigurationException, Authentication
from officeintegrator.src.com.zoho.officeintegrator.v1.create_document_parameters import CreateDocumentParameters
from officeintegrator.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations


class CreateDocument:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-writer-create-document.html
    @staticmethod
    def execute():
        CreateDocument.init_sdk()
        createDocumentParams = CreateDocumentParameters()

        # Optional Configuration - Add document meta in request to identify the file in Zoho Server
        documentInfo = DocumentInfo()
        documentInfo.set_document_name("New Document")
        documentInfo.set_document_id((round(time.time() * 1000)).__str__())

        createDocumentParams.set_document_info(documentInfo)

        # Optional Configuration - Add User meta in request to identify the user in document session
        userInfo = UserInfo()
        userInfo.set_user_id("1000")
        userInfo.set_display_name("User 1")

        createDocumentParams.set_user_info(userInfo)

        # Optional Configuration - Add callback settings to configure.
        # how file needs to be received while saving the document
        callbackSettings = CallbackSettings()

        # Optional Configuration - configure additional parameters
        # which can be received along with document while save callback
        saveUrlParams = {}

        saveUrlParams['id'] = '123131'
        saveUrlParams['auth_token'] = '1234'
        # Following $<> values will be replaced by actual value in callback request
        # To know more - https://www.zoho.com/officeintegrator/api/v1/zoho-writer-create-document.html#saveurl_params
        saveUrlParams['extension'] = '$format'
        saveUrlParams['document_name'] = '$filename'
        saveUrlParams['session_id'] = '$session_id'

        callbackSettings.set_save_url_params(saveUrlParams)

        # Optional Configuration - configure additional headers
        # which could be received in callback request headers while saving document
        saveUrlHeaders = {}

        saveUrlHeaders['access_token'] = '12dweds32r42wwds34'
        saveUrlHeaders['client_id'] = '12313111'

        callbackSettings.set_save_url_headers(saveUrlHeaders)

        callbackSettings.set_retries(1)
        callbackSettings.set_timeout(10000)
        callbackSettings.set_save_format("zdoc")
        callbackSettings.set_http_method_type("post")
        callbackSettings.set_save_url(
            "https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157123434d4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286")

        createDocumentParams.set_callback_settings(callbackSettings)

        # Optional Configuration - Set margin while creating document itself.
        # It's applicable only for new documents.
        margin = Margin()

        margin.set_top("1in")
        margin.set_bottom("1in")
        margin.set_left("1in")
        margin.set_right("1in")

        # Optional Configuration - Set default settings for document while creating document itself.
        # It's applicable only for new documents.
        documentDefaults = DocumentDefaults()

        documentDefaults.set_font_size(12)
        documentDefaults.set_paper_size("A4")
        documentDefaults.set_font_name("Arial")
        documentDefaults.set_track_changes("enabled")
        documentDefaults.set_orientation("landscape")

        documentDefaults.set_margin(margin)
        documentDefaults.set_language("ta")

        createDocumentParams.set_document_defaults(documentDefaults)

        # Optional Configuration
        editorSettings = EditorSettings()

        editorSettings.set_unit("in")
        editorSettings.set_language("en")
        editorSettings.set_view("pageview")

        createDocumentParams.set_editor_settings(editorSettings)

        # Optional Configuration
        uiOptions = UiOptions()

        uiOptions.set_dark_mode("show")
        uiOptions.set_file_menu("show")
        uiOptions.set_save_button("show")
        uiOptions.set_chat_panel("show")

        createDocumentParams.set_ui_options(uiOptions)

        # Optional Configuration - Configure permission values for session
        # based of you application requirement
        permissions = {}

        permissions["document.export"] = True
        permissions["document.print"] = True
        permissions["document.edit"] = True
        permissions["review.comment"] = True
        permissions["review.changes.resolve"] = True
        permissions["document.pausecollaboration"] = True
        permissions["document.fill"] = False

        createDocumentParams.set_permissions(permissions)

        v1Operations = V1Operations()
        response = v1Operations.create_document(createDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CreateDocumentResponse):
                    print('Document Id : ' + str(responseObject.get_document_id()))
                    print('Document Session ID : ' + str(responseObject.get_session_id()))
                    print('Document Session URL : ' + str(responseObject.get_document_url()))
                    print('Document Session Delete URL : ' + str(responseObject.get_session_delete_url()))
                    print('Document Delete URL : ' + str(responseObject.get_document_delete_url()))
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Document Creation Request Failed')

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


CreateDocument.execute()
