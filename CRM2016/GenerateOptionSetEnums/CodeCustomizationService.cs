// Define REMOVE_PROXY_TYPE_ASSEMBLY_ATTRIBUTE if you plan on compiling the output from
// this CrmSvcUtil extension with the output from the unextended CrmSvcUtil in the same
// assembly (this assembly attribute can only be defined once in the assembly).
#define REMOVE_PROXY_TYPE_ASSEMBLY_ATTRIBUTE

namespace GenerateOptionSetEnums
{
  using System;
  using System.Linq;
  using System.CodeDom;
  using Microsoft.Crm.Services.Utility;
  using System.Collections.Generic;
  using Microsoft.Xrm.Sdk.Metadata;
  using System.Globalization;
  using System.Text.RegularExpressions;

  public sealed class CodeCustomizationService : ICustomizeCodeDomService
  {
    /// <summary>
    /// Remove the unnecessary classes that we generated for entities. 
    /// </summary>
    public void CustomizeCodeDom(CodeCompileUnit codeUnit, IServiceProvider services)
    {
      var metadataProviderService = (IMetadataProviderService)services.GetService(typeof(IMetadataProviderService));
      IOrganizationMetadata metadata = metadataProviderService.LoadMetadata();
      string baseTypes = GetParameter("/basetypes");

      foreach (CodeNamespace codeNamespace in codeUnit.Namespaces)
      {
        foreach (CodeTypeDeclaration codeTypeDeclaration in codeNamespace.Types)
        {
          if (codeTypeDeclaration.IsClass)
          {
            codeTypeDeclaration.CustomAttributes.Clear();

            if (!string.IsNullOrEmpty(baseTypes))
            {
              codeTypeDeclaration.BaseTypes.Clear();
              codeTypeDeclaration.BaseTypes.Add(baseTypes);
            }

            foreach (EntityMetadata entityMetadata in metadata.Entities.Where(e => e.SchemaName == codeTypeDeclaration.Name))
            {
              for (var j = 0; j < codeTypeDeclaration.Members.Count; )
              {
                if (codeTypeDeclaration.Members[j].GetType() == typeof(System.CodeDom.CodeMemberProperty) && codeTypeDeclaration.Members[j].CustomAttributes.Count > 0 && codeTypeDeclaration.Members[j].CustomAttributes[0].AttributeType.BaseType == "Microsoft.Xrm.Sdk.AttributeLogicalNameAttribute")
                {
                  string attribute = ((System.CodeDom.CodePrimitiveExpression)codeTypeDeclaration.Members[j].CustomAttributes[0].Arguments[0].Value).Value.ToString();

                  AttributeMetadata attributeMetadata = entityMetadata.Attributes.Where(a => a.LogicalName == attribute).FirstOrDefault();
                  if (attributeMetadata != null)
                  {
                    IEnumerable<EnumItem> values = new List<EnumItem>();
                    switch (attributeMetadata.AttributeType.Value)
                    {
                      case AttributeTypeCode.Boolean: values = ToUniqueValues(((BooleanAttributeMetadata)attributeMetadata).OptionSet); break;
                      case AttributeTypeCode.Picklist: values = ToUniqueValues(((PicklistAttributeMetadata)attributeMetadata).OptionSet.Options); break;
                      case AttributeTypeCode.State: values = ToUniqueValues(((StateAttributeMetadata)attributeMetadata).OptionSet.Options); break;
                      case AttributeTypeCode.Status: values = ToUniqueValues(((StatusAttributeMetadata)attributeMetadata).OptionSet.Options); break;
                    }

                    switch (attributeMetadata.AttributeType.Value)
                    {
                      case AttributeTypeCode.Boolean:
                        CodeTypeDeclaration booleanClass = new CodeTypeDeclaration(string.Format("{0}Values", attributeMetadata.LogicalName));
                        booleanClass.IsClass = true;
                        booleanClass.Attributes = MemberAttributes.Public | MemberAttributes.Final;
                        booleanClass.CustomAttributes.Add(new CodeAttributeDeclaration { Name = "System.Xml.Serialization.XmlType", Arguments = { new CodeAttributeArgument { Name = "TypeName", Value = new CodePrimitiveExpression(string.Format("{0}.{1}.{2}Values", codeNamespace.Name, codeTypeDeclaration.Name, attributeMetadata.LogicalName)) } } });

                        foreach (EnumItem enumItem in values)
                        {
                          CodeMemberField codeMemberField = new CodeMemberField(typeof(bool), enumItem.Value);
                          codeMemberField.Attributes = MemberAttributes.Public | MemberAttributes.Const;
                          codeMemberField.InitExpression = new CodePrimitiveExpression(enumItem.Key != 0);
                          codeMemberField.CustomAttributes.Add(new CodeAttributeDeclaration { Name = "System.ComponentModel.Description", Arguments = { new CodeAttributeArgument { Value = new CodePrimitiveExpression(enumItem.Description) } } });
                          booleanClass.Members.Add(codeMemberField);
                        }

                        codeTypeDeclaration.Members[j] = booleanClass;
                        break;
                      case AttributeTypeCode.Picklist:
                      case AttributeTypeCode.State:
                      case AttributeTypeCode.Status:
                        CodeTypeDeclaration enumeration = new CodeTypeDeclaration
                        {
                          IsEnum = true,
                          Name = string.Format("{0}Values", attributeMetadata.LogicalName)
                        };
                        enumeration.CustomAttributes.Add(new CodeAttributeDeclaration { Name = "System.Xml.Serialization.XmlType", Arguments = { new CodeAttributeArgument { Name = "TypeName", Value = new CodePrimitiveExpression(string.Format("{0}.{1}.{2}Values", codeNamespace.Name, codeTypeDeclaration.Name, attributeMetadata.LogicalName)) } } });

                        foreach (EnumItem enumItem in values)
                        {
                          CodeMemberField codeMemberField = new CodeMemberField(string.Empty, enumItem.Value);
                          codeMemberField.InitExpression = new CodePrimitiveExpression(enumItem.Key);
                          codeMemberField.CustomAttributes.Add(new CodeAttributeDeclaration { Name = "System.ComponentModel.Description", Arguments = { new CodeAttributeArgument { Value = new CodePrimitiveExpression(enumItem.Description) } } });
                          enumeration.Members.Add(codeMemberField);
                        }

                        codeTypeDeclaration.Members[j] = enumeration;
                        break;
                    }
                    j++;
                  }
                  else
                  {
                    codeTypeDeclaration.Members.RemoveAt(j);
                  }
                }
                else
                {
                  codeTypeDeclaration.Members.RemoveAt(j);
                }
              }
            }
          }
        }
      }

#if REMOVE_PROXY_TYPE_ASSEMBLY_ATTRIBUTE
      foreach (CodeAttributeDeclaration attribute in codeUnit.AssemblyCustomAttributes)
      {
        if (attribute.AttributeType.BaseType == "Microsoft.Xrm.Sdk.Client.ProxyTypesAssemblyAttribute")
        {
          codeUnit.AssemblyCustomAttributes.Remove(attribute);
          break;
        }
      }
#endif
    }

    private static IEnumerable<EnumItem> ToUniqueValues(OptionMetadataCollection optionMetadataCollection)
    {
      return ToUniqueValues(optionMetadataCollection.Select(omc => new EnumItem(omc.Value.Value, omc.Label.UserLocalizedLabel.Label, omc.Label.UserLocalizedLabel.Label)).ToList());
    }

    private static IEnumerable<EnumItem> ToUniqueValues(BooleanOptionSetMetadata booleanOptionSetMetadata)
    {
      List<EnumItem> values = new List<EnumItem>();
      if (booleanOptionSetMetadata.FalseOption == null)
      {
        values.Add(new EnumItem(0, "No", "No"));
      }
      else
      {
        values.Add(new EnumItem(booleanOptionSetMetadata.FalseOption.Value.Value, booleanOptionSetMetadata.FalseOption.Label.UserLocalizedLabel.Label, booleanOptionSetMetadata.FalseOption.Label.UserLocalizedLabel.Label));
      }

      if (booleanOptionSetMetadata.TrueOption == null)
      {
        values.Add(new EnumItem(1, "Yes", "Yes"));
      }
      else
      {
        values.Add(new EnumItem(booleanOptionSetMetadata.TrueOption.Value.Value, booleanOptionSetMetadata.TrueOption.Label.UserLocalizedLabel.Label, booleanOptionSetMetadata.TrueOption.Label.UserLocalizedLabel.Label));
      }
      return ToUniqueValues(values);
    }

    private static IEnumerable<EnumItem> ToUniqueValues(IEnumerable<EnumItem> source)
    {
      Dictionary<string, int> uniqueValues = source.GroupBy(sv => sv.Value).ToDictionary(g => g.Key, g => g.Count());
      source.ToList().ForEach(sv => sv.Value = uniqueValues[sv.Value] == 1 ? sv.Value : string.Format("{0}_{1}", sv.Value, sv.Key.ToString(CultureInfo.InvariantCulture)));
      return source;
    }

    private static string ToValidIdentifier(string source)
    {
      string result = string.Empty;

      source = ReplaceSpecial(source ?? string.Empty);

      foreach (Match m in Regex.Matches(source, "[A-Za-z0-9]"))
      {
        result += m.ToString();
      }

      if (string.IsNullOrEmpty(source))
      {
        result = "None";
      }

      if (char.IsDigit(result[0]))
      {
        result = "_" + result;
      }

      return result;
    }

    private static string ReplaceSpecial(string value)
    {
      return value.Replace("<>", "NotEquals").Replace("!=", "NotEquals").Replace("=", "Equals").Replace("<", "LessThan").Replace(">", "GreaterThan").Replace("+", "Plus");
    }

    private class EnumItem
    {
      private int key;
      private string value;
      private string description;

      public EnumItem(int key, string value, string description)
      {
        this.key = key;
        this.value = ToValidIdentifier(value);
        this.description = description;
      }

      public int Key
      {
        get
        {
          return this.key;
        }
      }

      public string Value
      {
        get
        {
          return this.value;
        }
        set
        {
          this.value = value;
        }
      }

      public string Description
      {
        get
        {
          return this.description;
        }
      }
    }

    private static string GetParameter(string key)
    {
      string[] args = Environment.GetCommandLineArgs();
      foreach (string arg in args)
      {
        string[] argument = arg.Split(new char[] { ':' }, 2);
        if (argument.Length == 2 && argument[0].ToLowerInvariant() == key.ToLowerInvariant())
        {
          return argument[1].Trim(new char[] { '"' });
        }
      }

      return null;
    }
  }
}
