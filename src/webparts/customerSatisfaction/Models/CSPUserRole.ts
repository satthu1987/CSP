export interface ICSPUserRolePerson {
  Id: number;
  EMail: string;
  Title: string;
}

export interface ICSPUserRole {
  Id: number;
  Title?: string;
  PIC: ICSPUserRolePerson[];
  Role: string; // visitor | leader | manager | admin
}
