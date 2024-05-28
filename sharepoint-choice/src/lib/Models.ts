export interface UserQuery {
    queryParams: UserQueryParams;
}

export interface UserQueryParams {
    QueryString: string;
    MaximumEntitySuggestions: number;
    AllowEmailAddresses: boolean;
    AllowOnlyEmailAddresses: boolean;
    PrincipalType: number;
    PrincipalSource: number;
    SharePointGroupID: number;
}

export interface User {
    Key: string;
    Description: string;
    DisplayText: string;
    EntityType: string;
    ProviderDisplayName: string;
    ProviderName: string;
    IsResolved: boolean;
    EntityData: UserEntityData;
    MultipleMatches: any[];
}

export interface UserEntityData {
    IsAltSecIdPresent: string;
    Title: string;
    Email: string;
    MobilePhone: string;
    ObjectId: string;
    Department: string;
}