/**
 * AzureFileStorageRest
 * 
 * REST endpoint for the Outlook Add-in to browse and fetch Azure files.
 * This wraps your existing AzureFileStorageController methods behind a
 * REST API that the add-in can call with a Bearer token.
 *
 * Deploy this alongside your existing AzureFileStorageController class.
 *
 * Endpoints:
 *   POST /services/apexrest/AzureFileStorage/browseDirectory
 *        Body: { "directoryPath": "some/path" }
 *        Returns: JSON array of file/folder objects
 *
 *   POST /services/apexrest/AzureFileStorage/getFileContent
 *        Body: { "filePath": "some/path/file.pdf" }
 *        Returns: { "base64Content": "...", "contentType": "application/pdf" }
 *
 * Authentication:
 *   The Outlook add-in sends the Salesforce session token as a Bearer token.
 *   Salesforce REST API handles session validation automatically.
 *
 * CORS Note:
 *   Salesforce REST API does not support CORS from arbitrary origins.
 *   For the add-in to call this directly, you have two options:
 *     a) Use a Connected App + OAuth flow (recommended for production)
 *     b) Use a lightweight proxy/middleware (e.g., Heroku, Azure Function)
 *        that forwards requests to Salesforce with the session token
 *   See the README for setup details.
 */
@RestResource(urlMapping='/AzureFileStorage/*')
global with sharing class AzureFileStorageRest {

    @HttpPost
    global static void handlePost() {
        RestRequest req = RestContext.request;
        RestResponse res = RestContext.response;
        res.addHeader('Content-Type', 'application/json');

        try {
            // Extract method name from URL path
            // e.g., /services/apexrest/AzureFileStorage/browseDirectory
            String requestUri = req.requestURI;
            String methodName = requestUri.substringAfterLast('/');

            // Parse request body
            Map<String, Object> params = new Map<String, Object>();
            if (String.isNotBlank(req.requestBody?.toString())) {
                params = (Map<String, Object>) JSON.deserializeUntyped(
                    req.requestBody.toString()
                );
            }

            if (methodName == 'browseDirectory') {
                String directoryPath = (String) params.get('directoryPath');
                List<Object> result = AzureFileStorageController.browseDirectory(
                    directoryPath != null ? directoryPath : ''
                );
                res.responseBody = Blob.valueOf(JSON.serialize(result));
                res.statusCode = 200;

            } else if (methodName == 'getFileContent') {
                String filePath = (String) params.get('filePath');
                if (String.isBlank(filePath)) {
                    res.statusCode = 400;
                    res.responseBody = Blob.valueOf(
                        '{"error": "filePath parameter is required"}'
                    );
                    return;
                }
                Object result = AzureFileStorageController.getFileContent(filePath);
                res.responseBody = Blob.valueOf(JSON.serialize(result));
                res.statusCode = 200;

            } else {
                res.statusCode = 404;
                res.responseBody = Blob.valueOf(
                    '{"error": "Unknown method: ' + String.escapeSingleQuotes(methodName) + '"}'
                );
            }
        } catch (Exception ex) {
            res.statusCode = 500;
            res.responseBody = Blob.valueOf(
                '{"error": "' + String.escapeSingleQuotes(ex.getMessage()) + '"}'
            );
        }
    }
}
