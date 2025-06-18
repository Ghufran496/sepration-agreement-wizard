export interface PartyInfo {
  party1FirstName: string;
  party1MiddleName?: string;
  party1LastName: string;
  party2FirstName: string;
  party2MiddleName?: string;
  party2LastName: string;
  
  spousalSupportPayor?: string; // 'party1' or 'party2'
  spousalSupportRecipient?: string; // 'party1' or 'party2'
  childSupportPayor?: string; // 'party1' or 'party2'
  childSupportRecipient?: string; // 'party1' or 'party2'
  
  party1Role?: string; // 'Mother', 'Father', 'Wife', 'Husband'
  party2Role?: string; // 'Mother', 'Father', 'Wife', 'Husband'
  
  marriedDate?: Date;
  cohabitationDate?: Date;
  separationDate?: Date;
  
  children: Child[];
}

export interface Child {
  firstName: string;
  middleName?: string;
  lastName: string;
  birthdate: Date;
}
