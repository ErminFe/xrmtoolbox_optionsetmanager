using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace smartpoint.XrmToolBoxPlugins.OptionSetManager
{
  class Helper
  {
    public static EntityMetadata[] GetEntities(IOrganizationService organizationService)
    {
      Dictionary<string, string> attributesData = new Dictionary<string, string>();
      RetrieveAllEntitiesRequest metaDataRequest = new RetrieveAllEntitiesRequest();
      RetrieveAllEntitiesResponse metaDataResponse = new RetrieveAllEntitiesResponse();
      metaDataRequest.EntityFilters = EntityFilters.Entity;

      // Execute the request.
      metaDataResponse = (RetrieveAllEntitiesResponse)organizationService.Execute(metaDataRequest);
      var entities = metaDataResponse.EntityMetadata;

      return entities;
    }

    public static List<KeyValuePair<string, string>> GetEntyNames(IOrganizationService organizationService)
    {
      List<KeyValuePair<string, string>> result = new List<KeyValuePair<string, string>>();

      foreach (EntityMetadata entry in GetEntities(organizationService))
        result.Add(new KeyValuePair<string, string>(entry.LogicalName,
          entry.DisplayName.UserLocalizedLabel != null ? entry.DisplayName.UserLocalizedLabel.Label : entry.SchemaName));

      return result;
    }//GetEntyNames

    public static List<KeyValuePair<string, string>> GetOptionSetsForEntity(IOrganizationService organizationService, string logicalName)
    {
      List<KeyValuePair<string, string>> result = new List<KeyValuePair<string, string>>();

      RetrieveEntityRequest req = new RetrieveEntityRequest();
      req.EntityFilters = EntityFilters.Attributes;
      req.LogicalName = logicalName;
      RetrieveEntityResponse res = (RetrieveEntityResponse)organizationService.Execute(req);
      EntityMetadata metaData = res.EntityMetadata;
      foreach (AttributeMetadata attribute in metaData.Attributes)
      {
        if (attribute.AttributeType == AttributeTypeCode.Picklist)
          result.Add(new KeyValuePair<string, string>(attribute.LogicalName,
            attribute.DisplayName.UserLocalizedLabel != null ? attribute.DisplayName.UserLocalizedLabel.Label : attribute.SchemaName));
      }//foreach

      return result;
    }//GetEntyNames

    public static InsertOptionValueRequest CreateInsertOptionValueRequest(
      IOrganizationService organizationService,
      int userLanguageCode,
      string entityName,
      string attributeName,
      List<KeyValuePair<int, string>> labels,
      int value)
    {
      KeyValuePair<int, string> userLanguageLabel = labels.Where(_ => _.Key == userLanguageCode).FirstOrDefault();

      Label label = new Label(new LocalizedLabel(labels.Where(_ => _.Key == userLanguageCode).Select(_ => _.Value).FirstOrDefault(), userLanguageCode),
         labels.Where(_ => _.Key != userLanguageCode).Select(_ => new LocalizedLabel(_.Value, _.Key)).ToArray());

      // Create a request.
      return new InsertOptionValueRequest
      {
        AttributeLogicalName = attributeName,
        EntityLogicalName = entityName,
        Label = label,
        Value = value
      };

    }//CreateInsertOptionValueRequest

    public static UpdateOptionValueRequest CreateUpdateOptionValueRequest(
      IOrganizationService organizationService,
      int userLanguageCode,
      string entityName,
      string attributeName,
      List<KeyValuePair<int, string>> labels,
      int value)
    {
      KeyValuePair<int, string> userLanguageLabel = labels.Where(_ => _.Key == userLanguageCode).FirstOrDefault();

      Label label = new Label(new LocalizedLabel(labels.Where(_ => _.Key == userLanguageCode).Select(_ => _.Value).FirstOrDefault(), userLanguageCode),
         labels.Select(_ => new LocalizedLabel(_.Value, _.Key)).ToArray());

      // Create a request.
      return new UpdateOptionValueRequest
      {
        AttributeLogicalName = attributeName,
        EntityLogicalName = entityName,
        Label = label,
        Value = value,
        MergeLabels = true,

      };

    }//CreateInsertOptionValueRequest

    public static OptionMetadataCollection GetOptionSetEntries(IOrganizationService organizationService, string entityName, string attributeName)
    {
      RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest();
      retrieveAttributeRequest.EntityLogicalName = entityName;
      retrieveAttributeRequest.LogicalName = attributeName;
      retrieveAttributeRequest.RetrieveAsIfPublished = true;

      RetrieveAttributeResponse retrieveAttributeResponse =
        (RetrieveAttributeResponse)organizationService.Execute(retrieveAttributeRequest);
      PicklistAttributeMetadata picklistAttributeMetadata =
        (PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;

      return picklistAttributeMetadata.OptionSet.Options;
    }

  }
}
