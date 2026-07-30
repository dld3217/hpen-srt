import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';

const SITE_URL  = 'https://hpe.sharepoint.com/teams/EnterpriseSalesTeamHome';
const LIST_NAME = 'Contacts';

export interface IContact {
  id: number;
  name: string;
  email: string;
  category: string;
  businessUnit: string;
  geography: string;
  pocSolutions: string;
  phone: string;
}

export class ContactDirectoryService {
  constructor(private context: WebPartContext) {}

  async getAll(): Promise<IContact[]> {
    try {
      const sp = spfi(SITE_URL).using(SPFx(this.context));
      const items: { Id: number; Title: string; Email: string; Category: string; BusinessUnit: string; Geography: string; POCSolutions: string; Phone: string }[] =
        await sp.web.lists.getByTitle(LIST_NAME).items
          .select('Id', 'Title', 'Email', 'Category', 'BusinessUnit', 'Geography', 'POCSolutions', 'Phone')
          .orderBy('Title', true)
          .top(500)();
      return items.map(i => ({
        id: i.Id,
        name: i.Title || '',
        email: i.Email || '',
        category: i.Category || '',
        businessUnit: i.BusinessUnit || '',
        geography: i.Geography || '',
        pocSolutions: i.POCSolutions || '',
        phone: i.Phone || '',
      }));
    } catch {
      return [];
    }
  }

  suggestForSolutions(contacts: IContact[], solutionCodes: string[]): IContact[] {
    if (!solutionCodes.length) return [];
    return contacts.filter(c => {
      const cSols = c.pocSolutions.split(';').map(s => s.trim()).filter(Boolean);
      return solutionCodes.some(s => cSols.includes(s));
    });
  }
}
