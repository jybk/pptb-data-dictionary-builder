// Sample form, view and column metadata shared by the browser preview and the tests.
(function (root) {
    const L = (text) => ({ LocalizedLabels: [{ Label: text, LanguageCode: 1033 }], UserLocalizedLabel: { Label: text, LanguageCode: 1033 } });
    const lab = (en, nl) => `<labels><label description="${en}" languagecode="1033"/>${nl ? `<label description="${nl}" languagecode="1043"/>` : ""}</labels>`;
    const cell = (field, en, nl, extra = "") => `<cell id="{c-${field}}" ${extra}>${lab(en, nl)}<control id="${field}" classid="{4273EDBD-AC1D-40d3-9FB2-095C621B552D}" datafieldname="${field}" /></cell>`;

    const formXml = `<form>
  <header id="{h}"><rows><row>
    ${cell("ownerid", "Owner", "Eigenaar")}
    ${cell("statuscode", "Status Reason", "Statusreden")}
  </row></rows></header>
  <tabs>
    <tab name="SUMMARY_TAB" id="{t1}">${lab("Summary", "Overzicht")}
      <columns>
        <column width="50%"><sections>
          <section name="ACCOUNT_INFORMATION" id="{s1}">${lab("Account Information", "Accountgegevens")}<rows>
            <row>${cell("name", "Account Name", "Accountnaam")}</row>
            <row>${cell("telephone1", "Phone", "Telefoon")}</row>
            <row>${cell("websiteurl", "Website", "Website")}</row>
            <row>${cell("parentaccountid", "Parent Account", "Bovenliggend account")}</row>
            <row><cell id="{spacer}"></cell></row>
          </rows></section>
          <section name="ADDRESS" id="{s2}">${lab("Address", "Adres")}<rows>
            <row>${cell("address1_composite", "Address 1", "Adres 1")}</row>
          </rows></section>
        </sections></column>
        <column width="50%"><sections>
          <section name="CONTACTS" id="{s3}">${lab("Contacts", "Contactpersonen")}<rows>
            <row>${cell("primarycontactid", "Primary Contact", "Primaire contactpersoon")}</row>
            <row><cell id="{c-grid}">${lab("Contacts", "Contactpersonen")}<control id="Contacts" classid="{E7A81278-8635-4d9e-8D4D-59480B391C5B}"><parameters><TargetEntityType>contact</TargetEntityType><RelationshipName>contact_customer_accounts</RelationshipName></parameters></control></cell></row>
          </rows></section>
        </sections></column>
      </columns>
    </tab>
    <tab name="DETAILS_TAB" id="{t2}">${lab("Details", "Details")}
      <columns><column width="100%"><sections>
        <section name="COMPANY_PROFILE" id="{s4}">${lab("Company Profile", "Bedrijfsprofiel")}<rows>
          <row>${cell("industrycode", "Industry", "Branche")}</row>
          <row>${cell("revenue", "Annual Revenue", "Jaaromzet")}</row>
          <row>${cell("numberofemployees", "No. of Employees", "Aantal werknemers")}</row>
          <row>${cell("name", "Account Name", "Accountnaam")}</row>
          <row>${cell("new_onboardedon", "Onboarded On", "Onboarding op")}</row>
        </rows></section>
        <section name="PREFERENCES" id="{s5}">${lab("Contact Preferences", "Contactvoorkeuren")}<rows>
          <row>${cell("donotemail", "Email", "E-mail")}</row>
          <row>${cell("new_internalscore", "Internal Score", "Interne score", 'visible="false"')}</row>
          <row>${cell("description", "Description", "Beschrijving")}</row>
        </rows></section>
      </sections></column></columns>
    </tab>
    <tab name="ADMIN_TAB" id="{t3}" visible="false">${lab("Administration", "Beheer")}
      <columns><column width="100%"><sections>
        <section name="ADMIN" id="{s6}">${lab("Administration", "Beheer")}<rows>
          <row>${cell("createdon", "Created On", "Gemaakt op")}</row>
        </rows></section>
      </sections></column></columns>
    </tab>
  </tabs>
</form>`;

    const base = (logical, schema, display, type, extra = {}) => ({
        LogicalName: logical,
        SchemaName: schema,
        DisplayName: L(display),
        Description: L(extra.desc ?? ""),
        AttributeType: type.replace(/Type$/, ""),
        AttributeTypeName: { Value: type },
        RequiredLevel: { Value: extra.req ?? "None" },
        IsCustomAttribute: logical.startsWith("new_"),
        IsValidForCreate: true,
        IsValidForUpdate: true,
        IsValidForAdvancedFind: { Value: true },
        IsAuditEnabled: { Value: true },
        IsSecured: false,
        SourceType: 0,
        ...extra.props,
    });

    const accountAttributes = [
        base("name", "Name", "Account Name", "StringType", { req: "ApplicationRequired", desc: "Type the company or business name.", props: { MaxLength: 160, FormatName: { Value: "Text" } } }),
        base("telephone1", "Telephone1", "Main Phone", "StringType", { desc: "Type the main phone number for this account.", props: { MaxLength: 50, FormatName: { Value: "Phone" } } }),
        base("websiteurl", "WebSiteURL", "Website", "StringType", { desc: "Type the account's website URL.", props: { MaxLength: 200, FormatName: { Value: "Url" } } }),
        base("parentaccountid", "ParentAccountId", "Parent Account", "LookupType", { desc: "Choose the parent account to show parent and child businesses in reporting.", props: { Targets: ["account"] } }),
        base("address1_composite", "Address1_Composite", "Address 1", "MemoType", { desc: "Shows the complete primary address.", props: { MaxLength: 1000, IsValidForCreate: false, IsValidForUpdate: false } }),
        base("primarycontactid", "PrimaryContactId", "Primary Contact", "LookupType", { req: "Recommended", desc: "Choose the primary contact for the account.", props: { Targets: ["contact"] } }),
        base("industrycode", "IndustryCode", "Industry", "PicklistType", { desc: "Select the account's primary industry for use in segmentation.", props: { DefaultFormValue: -1 } }),
        base("revenue", "Revenue", "Annual Revenue", "MoneyType", { desc: "Type the annual revenue for the account.", props: { MinValue: 0, MaxValue: 100000000000000, Precision: 2 } }),
        base("numberofemployees", "NumberOfEmployees", "Number of Employees", "IntegerType", { desc: "Type the number of employees that work at the account.", props: { MinValue: 0, MaxValue: 1000000000, Format: "None" } }),
        base("new_onboardedon", "new_OnboardedOn", "Onboarded On", "DateTimeType", { desc: "Date the customer completed onboarding.", props: { Format: "DateOnly", DateTimeBehavior: { Value: "DateOnly" } } }),
        base("donotemail", "DoNotEMail", "Do not allow Emails", "BooleanType", { desc: "Select whether the account allows direct email.", props: { DefaultValue: false } }),
        base("new_internalscore", "new_InternalScore", "Internal Score", "DecimalType", { desc: "Score maintained by the account team.", props: { MinValue: 0, MaxValue: 100, Precision: 1, SourceType: 1 } }),
        base("description", "Description", "Description", "MemoType", { desc: "Type additional information to describe the account.", props: { MaxLength: 2000 } }),
        base("createdon", "CreatedOn", "Created On", "DateTimeType", { props: { Format: "DateAndTime", DateTimeBehavior: { Value: "UserLocal" }, IsValidForCreate: false, IsValidForUpdate: false } }),
        base("ownerid", "OwnerId", "Owner", "OwnerType", { req: "SystemRequired", desc: "Enter the user or team who is assigned to manage the record.", props: { Targets: ["systemuser", "team"] } }),
        base("statuscode", "StatusCode", "Status Reason", "StatusType", { desc: "Select the account's status.", props: { DefaultFormValue: 1 } }),
    ];
    const opt = (value, label) => ({ Value: value, Label: L(label) });
    const accountOptionSets = {
        Picklist: [{ LogicalName: "industrycode", OptionSet: { Options: [opt(1, "Accounting"), opt(2, "Agriculture"), opt(3, "Broadcasting"), opt(4, "Brokers"), opt(5, "Construction"), opt(6, "Consulting"), opt(7, "Consumer Services"), opt(8, "Design and Creative")] } }],
        Status: [{ LogicalName: "statuscode", OptionSet: { Options: [opt(1, "Active"), opt(2, "Inactive")] } }],
        Boolean: [{ LogicalName: "donotemail", OptionSet: { TrueOption: opt(1, "Do Not Allow"), FalseOption: opt(0, "Allow") } }],
    };
    const contactAttributes = [base("emailaddress1", "EMailAddress1", "Email", "StringType", { desc: "Type the primary email address for the contact.", props: { MaxLength: 100, FormatName: { Value: "Email" } } })];

    const layoutXml = `<grid name="resultset" object="1" jump="name" select="1" icon="1" preview="1"><row name="result" id="accountid"><cell name="name" width="300"/><cell name="telephone1" width="100"/><cell name="pc.emailaddress1" width="150"/><cell name="zz.unknowncol" width="100"/></row></grid>`;
    const fetchXml = `<fetch><entity name="account"><attribute name="name"/><link-entity name="contact" from="contactid" to="primarycontactid" alias="pc" link-type="outer"><attribute name="emailaddress1"/></link-entity></entity></fetch>`;

    root.DDB_FIXTURES = { formXml, layoutXml, fetchXml, accountAttributes, accountOptionSets, contactAttributes, L };
})(typeof window !== "undefined" ? window : globalThis);
