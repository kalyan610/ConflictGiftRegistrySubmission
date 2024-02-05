import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/files";


export default class Service {

    public mysitecontext: any;

    public constructor(siteUrl: string, Sitecontext: any) {
        this.mysitecontext = Sitecontext;


        sp.setup({
            sp: {
                baseUrl: siteUrl

            },
        });

    }

public async test123():Promise<any>
{

    this.Save('','','','','','','','','','','','','','','','','','','','','','','');

}


    public async GetAllBussinessUnits():Promise<any>
    {
 
     return await sp.web.lists.getByTitle("BusinessUnits").items.select('Title','ID').expand().get().then(function (data:any) {
 
     return data;
 
 
     });
 
 
    }

    public async GetAllgiftregistries():Promise<any>
    {
 
     return await sp.web.lists.getByTitle("GiftRegistry").items.select('Title','ID','RequestType').expand().get().then(function (data:any) {
 
     return data;
 
 
     });
 
 
    }

    public async GetAllCurrencies():Promise<any>
    {
 
     return await sp.web.lists.getByTitle("Currency").items.select('Title','ID').expand().get().then(function (data:any) {
 
     return data;
 
 
     });
 
 
    }

    public async GetCompanyType():Promise<any>
    {
 
     return await sp.web.lists.getByTitle("CompanyType").items.select('Title','ID').expand().get().then(function (data:any) {
 
     return data;
 
 
     });
 
 
    }

    public async GetFinancialServices():Promise<any>
    {
 
     return await sp.web.lists.getByTitle("FinancialServices").items.select('Title','ID').expand().get().then(function (data:any) {
 
     return data;
 
 
     });
 
 
    }

     public async getEnvironment():Promise<any>
     {
     
     return await sp.web.lists.getByTitle("Environment").items.select('Title','ID').expand().get().then(function (data:any) {
      
     return data;
     
     });
    }     

    public async getUserByLogin(LoginName:string):Promise<any>{
        try{
            const user = await sp.web.siteUsers.getByLoginName(LoginName).get();
            return user;
        }catch(error){
            console.log(error);
        }
    }

    public async getCurrentUserSiteGroups(): Promise<any[]> {

        try {

            return (await sp.web.currentUser.groups.select("Id,Title,Description,OwnerTitle,OnlyAllowMembersViewMembership,AllowMembersEditMembership,Owner/Id,Owner/LoginName").expand('Owner').get());

        }
        catch {
            throw 'get current user site groups failed.';
        }

    }

    public async Save (
        YourName:string,
        MyBussinessUnitVal:string,

        //WorkEmployment

        DoYouwishtoRegisterYesorNo:string,
        MyCompanyName:string,
        MyTypeofcompregisternumber:string,
        MyPosition:string,
        MyDetails:string,
         //PaidCheckbox
         MypaidorUnpaid:string,
         MyTypeofCompany:string,
         MyCompanyInvolvedFinancialServices:string,
         MyBoardofDirectoryYesorNo:string,
         MyBoardofDirectortCheckbox:string,
         //End
         MyStockValue:string,
         //InsiderInformation
         MyinsideryesorNo:string,
         MyyourFullName:string,
         Mydtinsider:string,
         Myclientname:string,
         Myclientproject:string,
         Mywhomesclated:string,
         Mydtpublic:string,
         //End

         //Conflict
        MyConflictYesNo:string,
        MyDetailsofConflict:string,
        MyBussinessString:string
        //End

        


      )  {

        let Myval='Completed';
    
        try
        {

         let Varmyval= await sp.web.lists.getByTitle("ConflictRegistrySubmissions").items.add({
    
            Name:YourName,
            BussinessUnitId:MyBussinessUnitVal,
            DoYouwishtoRegister:DoYouwishtoRegisterYesorNo,
            CompanyName:MyCompanyName,
            TypeOfcompany:MyTypeofcompregisternumber,
            Position:MyPosition,
            Details:MyDetails,
            PaidorUnpaid:MypaidorUnpaid,
            ActiveorDormamantCompany:MyTypeofCompany,
            CompanyInvolvedFinancialServices:MyCompanyInvolvedFinancialServices,
            SitonBoardoFDirectors:MyBoardofDirectoryYesorNo,
            OBoardCapcoClient:MyBoardofDirectortCheckbox,
            //Shareownership
            CompanyStock:MyStockValue,
            //End
            //InsiderInformation

            InsiderTraderAdded:MyinsideryesorNo,
            YourFullName:MyyourFullName,
            SelectDate:Mydtinsider,
            Clientname:Myclientname,
            Projectname:Myclientproject,
            WhomEsclated:Mywhomesclated,
            Anticpateddate:Mydtpublic,
            //End

            //Confict

            ConflictofIntrest:MyConflictYesNo,
            DetailsofConflict:MyDetailsofConflict,

            //End

            //Last

            BussinessValue:MyBussinessString,
            //End

            Title: "Request Created",
            
        }).then (async r => {
          // this will add an attachment to the item we just created to push t sharepoint list
    
          return Myval;
    
    
        })
    
    
        return Varmyval;
    
        
      }
    
    
    
      catch (error) {
        console.log(error);
      }
    
    
      
     }

     public async SaveAllNOS (
        YourName:string,
        MyBussinessUnitVal:string,

        //WorkEmployment

        DoYouwishtoRegisterYesorNo:string,
        //End
         MyStockValue:string,
         //InsiderInformation
         MyinsideryesorNo:string,
         //End

         //Conflict
        MyConflictYesNo:string,
        
        //End

        MyBussinessString:string


      )  {

        let Myval='Completed';
    
        try
        {

         let Varmyval= await sp.web.lists.getByTitle("ConflictRegistrySubmissions").items.add({
    
            Name:YourName,
            BussinessUnitId:MyBussinessUnitVal,
            DoYouwishtoRegister:DoYouwishtoRegisterYesorNo,
           //Shareownership
            CompanyStock:MyStockValue,
            //End
            //InsiderInformation
            InsiderTraderAdded:MyinsideryesorNo,
            //End
            //Confict
            ConflictofIntrest:MyConflictYesNo,
           //End

           BussinessValue:MyBussinessString,

            Title: "Request With All NOS",
            
            
        }).then (async r => {
          // this will add an attachment to the item we just created to push t sharepoint list
    
          return Myval;
    
    
        })
    
    
        return Varmyval;
    
        
      }
    
    
    
      catch (error) {
        console.log(error);
      }
    
    
      
     }



}