---
page_type: sample
languages:
- csharp
products:
- dotnet
- ms-graph
description: "Describes how you can batch request submitted to Microsoft Graph."
urlFragment: "update-this-to-unique-url-stub"
---

# Batching requests to Microsoft Graph

When using Microsoft Graph to manage inboxes and calendars you might end up needing to submit multiple request in one go for performance and efficiency reasons. Microsoft Graph has batching support for this but it might work a little different than you would expect.

Basically you can optimize the overhead of submitting multiple reperate request by turning them into a single batch payload but internally Microsoft Graph still handles each request reflected in the payload seperately. This means you have to walk through the response to see whether each request succeeded as this is not an atomic transaction. Below is a description of how a batch can be constructed and handled at runtime.

We setup a batch of requests by creating a BatchRequestContent object and inserting a series of BatchRequestStep objects which contain the HttpRequestMessage object that holds a request object that is exactly the same as the ones you would submit seperately.

```
 List<BatchRequestContent> batches = new List<BatchRequestContent>();
 
 var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, $"https://graph.microsoft.com/v1.0/users/{config["CalendarEmail"]}/events")
                {
                    Content = new StringContent(JsonConvert.SerializeObject(e), Encoding.UTF8, "application/json")
                };

BatchRequestStep requestStep = new BatchRequestStep(events.IndexOf(e).ToString(), httpRequestMessage, null);
batchRequestContent.AddBatchRequestStep(requestStep);
 ```
Next we submit the batch to Microsoft Graph which under the hood calls out to the https://graph.microsoft.com/v1.0/$batch endpoint with a payload that combines all requests in this format:

```
"requests": [
    {
      "id": "1",
      "method": "GET",
      "url": "/me/drive/root:/{file}:/content"
    },
    ...
```
The ID is important as we need to track the status of these request and because the response is unordered.

After the post we call GetResponsesAsync() to get a dictionary to walk through. We can call each individual response by using the Id (dictionary key).


```
response = await client.Batch.Request().PostAsync(batch);
Dictionary<string, HttpResponseMessage> responses = await response.GetResponsesAsync();

   foreach (string key in responses.Keys)
                {
                    HttpResponseMessage httpResponse = await response.GetResponseByIdAsync(key);
                    var responseContent = await httpResponse.Content.ReadAsStringAsync();

                   ...
```

The response clearly show us the result come in randomly.

![alt text](https://github.com/valeryjacobs/MSGraphBatch/blob/master/UnorderedBatchResponse.PNG "Logo Title Text 1")


## Prerequisites & setup

To run the sample you need an O365 tenant and a (dev) user account which calendar will be used to add events in batches and remove them afterwards. Additionally an application registration needs to be added (or reused) in the Azure AD directory belonging to the O365 tenant in use, including a client secret as we are using the client credential flow to aquire an accesstoken for Microsoft Graph.

Add a local.settings.json to the project with these settings and update their values according to your tenant and app registration setup.

```
{
  "Authority": "https://login.microsoftonline.com/",
  "ClientId": "[Add the client id (application id) from the application registration in Azure AD]",
  "ClientSecret": "[Add the client secret from the application registration in Azure AD]",
  "DateFormat": "yyyy-MM-ddTHH:mm:ss",
  "Scope": "https://graph.microsoft.com/.default",
  "TenantId": "[Azure AD tenant ID]",
  "CalendarEmail": "[Email address of the dev user's calendar in which objects will be created and deleted]"

}
```

## Running the sample

Run the application and keep the console window and a browser open with the calendar view of the respective user.

## Key concepts & considerations

Microsoft Graph throttles request as described in [these docs](https://docs.microsoft.com/en-us/graph/throttling#outlook-service-limits). When the limits are exceeded you should expect to see HTTP 429 responses that requires us to back of and retry at a later point in time.

Additionally the amount of steps in a batch is currently limited to ony 20 items (!). To submit more requests we unfortunately need to 'batch the batches'. The sample code shows a simple approach to do this.

## Contributing

This little sample project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
