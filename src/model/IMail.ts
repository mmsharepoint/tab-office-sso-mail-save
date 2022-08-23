export interface IMail {
  id: string;
  from: string;
  subject: string;
  hasAttachments: boolean;
  receivedDateTime: string;
}