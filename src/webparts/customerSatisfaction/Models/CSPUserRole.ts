export interface ICSPUserRole {
  Id: number;
  Title?: string;
  PIC: {
    Id: number;
    EMail: string;
    Title: string;
  };
  Role: string; // visitor | leader | manager | admin
}
