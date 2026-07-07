export interface IActionplan {
  Id: number;
  Title: string;
  Service: string;
  CustomerFeedback: string[];
  Actions: string[];
  PIC: {
    Title: string;
    EMail: string;
  };
  Timeline: string;
  Status: string;
  Results: string[];
  RelatedLinks: string;
  Year: string;
  Category: string;
  ProductLine: string;
  Department: string;
  Division: string;
}
