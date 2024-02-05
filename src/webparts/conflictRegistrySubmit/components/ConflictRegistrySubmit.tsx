import * as React from 'react';
import styles from './ConflictRegistrySubmit.module.scss';
import { IConflictRegistrySubmitProps } from './IConflictRegistrySubmitProps';

import { sp } from "@pnp/sp/presets/all";


import {PrimaryButton } from 'office-ui-fabric-react/lib/Button';

import { Dropdown,IDropdownStyles,IDropdownOption} from 'office-ui-fabric-react/lib/Dropdown';
import {Stack,StackItem,IStackStyles} from 'office-ui-fabric-react'; 
// import Collapsible from 'react-collapsible'

import { Checkbox} from 'office-ui-fabric-react/lib/Checkbox';

import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';

import Service from './Service1';
// import { each } from '@microsoft/sp-lodash-subset';

// const stackTokens = { childrenGap: 50 };
// const stackStyles: Partial<IStackStyles> = { root: { padding: 10} };

let paidArray:any=[];

let BoardofDirectorsArray:any=[];

let paidorunpaidCheckboxvalue='';

let BoardofDirectorsCheckboxValue='';
//let FinalPaidArray='';

let Envval='';

let BussinessValuetext='';

const stackTokens1 = { childrenGap: 80 };
const drpYesorNo:IDropdownOption[]=[  { key: "Yes", text: "Yes"},  { key: "No", text: "No" }];  

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 350 },
};


const stackButtonStyles: Partial<IStackStyles> = { root: { width: 20 } };

const dropdownStyles1: Partial<IDropdownStyles> = {
  dropdown: { width: 350 },
};

//let GivenchoiceValues: string[]=['Apple','Banana'];


export interface IConflictRegistrationSubmit
{
  
  yourName:any;
  BussinessUnitItems:any;
  MyBussinessUnitValue:any;
  MyEmployeeWorkYesorNo:boolean;
  MyEmploymentYesNo:any;
  Contactname:any;
  Companyname:any;
  RegisterNumber:any;
  Position:any;
  details:any;
  choiceValues: string[];
  choiceValues1: string[];
  choiceValuesonBoarddirectors: string[];
  CompanyTypes:any;
  ComanyFinancialServiceTypes:any;
  MyComanyValue:any;
  MyCompanyText:any;

  CompanyFinancialServices:any;
  ComanyFinancialService:any;
  MyFinancialService:any;
  MyFinancialServiceText:any;
  MyBoardofDirectorsYesorNo:any;

  Myboardofdirecrosboolean:boolean;
  MyWorkEmploymentboolean:boolean;

  StockValue:any;
  insideryesorNo:any;
  insideryesorNoboolean:boolean;

  yourFullName:any;
  dtinsider:any;
  clientname:any;
  clientproject:any;
  whomesclated:any;
  dtpublic:any;

  MyConflictYesNo:any;
  MyConflictbollean:boolean;
  detailsconflict:any;

  MyDornamantKey:any;
  MyDornamantValue:any;



}

export default class ConflictRegistrySubmit extends React.Component<IConflictRegistrySubmitProps, IConflictRegistrationSubmit> {
  
  public _service: any;
  public GlobalService1: any;
  protected ppl:any;
  
  public constructor(props:IConflictRegistrySubmitProps) {

    super(props);

    this.state={

      
      yourName:"",
      BussinessUnitItems:[],
      MyBussinessUnitValue:"",
      MyEmployeeWorkYesorNo:false,
      MyEmploymentYesNo:"",
      Contactname:"",
      Companyname:"",
      RegisterNumber:"",
      Position:"",
      details:"",
      choiceValues: [],
      choiceValues1: [],
      choiceValuesonBoarddirectors:[],
      CompanyTypes:[],
      ComanyFinancialServiceTypes:[],
      MyComanyValue:"",
      MyCompanyText:"",
      CompanyFinancialServices:[],
      ComanyFinancialService:"",
      MyFinancialService:"",
      MyFinancialServiceText:"",
      MyBoardofDirectorsYesorNo:"",
      Myboardofdirecrosboolean:false,
      MyWorkEmploymentboolean:false,
      StockValue:"",
      insideryesorNo:"",
      insideryesorNoboolean:false,
      yourFullName:"",
      dtinsider:"",
      clientname:"",
      clientproject:"",
      whomesclated:"",
      dtpublic:"",
      MyConflictYesNo:"",
      MyConflictbollean:false,
      detailsconflict:"",
      MyDornamantKey:"",
      MyDornamantValue:""
      

    

      

    };

    this._service = new Service(this.props.url, this.props.context);

    this.GetAllBussinessUnits();
    this.GetAllCompanyType();
    this.GetAllFinancialServices();
    this.GetEnvironment();

    
  }

  private _onFormatDate1 = (date: Date): string => {
    return date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear();
  
  
  };
  

  
  private changeContactName(data: any): void {

    this.setState({ Contactname: data.target.value });

  }

  private hadleBussinessUnit(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {

    
    this.setState({ MyBussinessUnitValue:item.key });

    BussinessValuetext=item.text;

  }

  public async GetAllBussinessUnits() {

   
    
    var data = await this._service.GetAllBussinessUnits();

    console.log(data);

    var AllBussinessUnits: any = [];

    for (var k in data) {

      AllBussinessUnits.push({ key: data[k].ID, text: data[k].Title});
    }

    console.log(AllBussinessUnits);

    
   this.setState({ BussinessUnitItems: AllBussinessUnits });

   
  

  }

  private handleEmploymentYesNo(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {

    
    this.setState({ MyEmploymentYesNo:item.key });

    if(item.key=='Yes')
    {

      this.setState({ MyWorkEmploymentboolean:true})
    }

    else
    {

      this.setState({ MyWorkEmploymentboolean:false})

    }

    
  }

  private handleBoardofDirectorsYesNo(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {

    this.setState({ MyBoardofDirectorsYesorNo:item.key });


    if(item.key=='Yes')
    {

      this.setState({ Myboardofdirecrosboolean:true})

    }

    else
    {

      this.setState({ Myboardofdirecrosboolean:false})

    }

    
  }

  private handleConflictsYesNo(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {

    this.setState({ MyConflictYesNo:item.key });


    if(item.key=='Yes')
    {

      this.setState({ MyConflictbollean:true})

    }

    else
    {

      this.setState({ MyConflictbollean:false})

    }

    
  }

  private changeCompanyName(data: any): void {

    this.setState({ Companyname: data.target.value });

  }

  private changeRegisterNumber(data: any): void {

    this.setState({ RegisterNumber: data.target.value });

  }

  private changePosition(data: any): void {

    this.setState({ Position: data.target.value });

  }

  private changedetails(data: any): void {

    this.setState({ details: data.target.value });

  }

  private changedetailsconflict(data: any): void {

    this.setState({ detailsconflict: data.target.value });

  }


  public async componentDidMount() {
    try {


      const choiceValues: string[] = await this.getCheckboxChoices("CLEARANCELIST");
      const choiceValues1: string[] = await this.getCheckboxChoicesboardofDirectors("BoardDetails");
      
   
      this.setState({ choiceValues});
      this.setState({ choiceValues1});
      
    } catch (error) {
      console.error(error);
    }
  }

  public async GetAllCompanyType() {

   
    
    var data = await this._service.GetCompanyType();

    console.log(data);

    var AllCompanyTypes: any = [];

    for (var k in data) {

      AllCompanyTypes.push({ key: data[k].ID, text: data[k].Title});
    }

    console.log(AllCompanyTypes);

    
   this.setState({ CompanyTypes: AllCompanyTypes });
  

  }

  public async GetAllFinancialServices() {

   
    
    var data = await this._service.GetFinancialServices();

    console.log(data);

    var AllFinancialServices: any = [];

    for (var k in data) {

      AllFinancialServices.push({ key: data[k].ID, text: data[k].Title});
    }

    console.log(AllFinancialServices);

    
   this.setState({ ComanyFinancialServiceTypes: AllFinancialServices });
  

  }

  public getCheckboxChoices = (fieldname: string): Promise<string[]> => {
    return new Promise<string[]>((resolve, reject) => {
      try {
        sp.web.lists.getByTitle("TypesofItems").fields
          .getByInternalNameOrTitle(fieldname)
          .select('Choices')
          .get()
          .then((field:any) => {
            const multivaluechoices: string[] = [];
            for (var i = 0; i < field["Choices"].length; i++) {
              multivaluechoices.push(field["Choices"][i]?.toString());
            }
            resolve(multivaluechoices);
          });
      } catch (error) {
        console.error(error);
        reject(error);
      }
    })
  }

  public getCheckboxChoicesboardofDirectors = (fieldname: string): Promise<string[]> => {
    return new Promise<string[]>((resolve, reject) => {
      try {
        sp.web.lists.getByTitle("BoardofDirector").fields
          .getByInternalNameOrTitle(fieldname)
          .select('Choices')
          .get()
          .then((field:any) => {
            const multivaluechoices: string[] = [];
            for (var i = 0; i < field["Choices"].length; i++) {
              multivaluechoices.push(field["Choices"][i]?.toString());
            }
            resolve(multivaluechoices);
          });
      } catch (error) {
        console.error(error);
        reject(error);
      }
    })
  }

  private _onCheckboxEmployment = async (ev: React.FormEvent<HTMLElement>, isChecked: boolean): Promise<void> => {
    try {
      //let updtedSelectedChoiceValues: string[]=[...this.state.choiceValues];

      let updtedSelectedChoiceValues: string[]=[];

      //let str = '';

      const choiceValue: string = ev.currentTarget["ariaLabel"];
      if (isChecked) {
        
        updtedSelectedChoiceValues.push(choiceValue);
        //this.setState({PaidArray:choiceValue});
        paidArray.push(choiceValue);
      } else {
        updtedSelectedChoiceValues = this.state.choiceValues.filter(sc => sc !== choiceValue);
      }

      console.log(updtedSelectedChoiceValues);
      console.log(paidArray.length);

    for(let count=0;count<paidArray.length;count++)
     {

      paidorunpaidCheckboxvalue = paidArray.toString();
      
    }

      

 
      
    } catch (error) {
      console.error(error);
    }
  }

  
  private _onCheckboxBoradofDirectors = async (ev: React.FormEvent<HTMLElement>, isChecked: boolean): Promise<void> => {
    try {

      let updtedSelectedChoiceValuesonBoard: string[]=[...this.state.choiceValuesonBoarddirectors];

      const choiceValue1: string = ev.currentTarget["ariaLabel"];
      if (isChecked) {
        updtedSelectedChoiceValuesonBoard.push(choiceValue1);
        BoardofDirectorsArray.push(choiceValue1);
      } else {
        updtedSelectedChoiceValuesonBoard = this.state.choiceValuesonBoarddirectors.filter(sc => sc !== choiceValue1);
      }


      for(let count=0;count<BoardofDirectorsArray.length;count++)
      {
 
       BoardofDirectorsCheckboxValue = BoardofDirectorsArray.toString();
       
     }

      
       
      
    } catch (error) {
      console.error(error);
    }
  }

  private handleActiveorDornamant(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {

    
    this.setState({ MyDornamantKey:item.key });
 
    this.setState({MyDornamantValue:item.text});


    
  }

  private handlefinancialservice(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {

    
    this.setState({ MyFinancialService:item.key });
 
    this.setState({MyFinancialServiceText:item.text});


    
  }


  public async GetEnvironment()
  {

    var data = await this._service.getEnvironment();

    console.log(data);

    var AllEnvironments: any = [];

    for (var k in data) {

      AllEnvironments.push({ key: data[k].ID, text: data[k].Title});

      Envval=data[0].Title;

      
    }

   
  }


  private changeStockvalue(data: any): void {

    //const re = /^[0-9\b]+$/;
      //if (data.target.value === '' || re.test(data.target.value)) {
        this.setState({ StockValue: data.target.value });
      //}

  }

  private handleInsiderYesorNo(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption):void {

    this.setState({ insideryesorNo:item.key });

    if(item.key=='Yes')
    {

      this.setState({ insideryesorNoboolean:true})
    }

    else
    {

      this.setState({ insideryesorNoboolean:false})

    }
  }

  private changeFullName(data: any): void {

    this.setState({ yourFullName: data.target.value });

  }

  private changeClientname(data: any): void {

    this.setState({ clientname: data.target.value });

  }

  private changeProjectName(data: any): void {

    this.setState({ clientproject: data.target.value });

  }
  private changewhotoesclated(data: any): void {

    this.setState({ whomesclated: data.target.value });

  }

   public handldtInsiderDateChange = (date: any) => {

    this.setState({ dtinsider: date });



   }

   public handldtPublicDateChange = (date: any) => {

    this.setState({ dtpublic: date });

   }

   private OnBtnClickSubmit():void{

    //Section1

    if(this.state.Contactname=="")
    {

      alert('Please enter your name')
    }

    else if(this.state.MyBussinessUnitValue=='')
    {

      alert('Please select the business unit')
    }

    //EndSection1

    
    else if(this.state.MyEmploymentYesNo=="")
    {
      alert('Please select the work Employment')
    }

    //WorkEmployment

    if(this.state.MyEmploymentYesNo=='Yes')
    {

      if(this.state.Companyname=="")
      {

        alert('Please enter the CompanyName')

      }

      else if(this.state.RegisterNumber=="")
      {

        alert('Please enter the RegisterNumber')

      }

      else if(this.state.Position=="")
      {

        alert('Please enter the Position')
      }

      else if(this.state.details=="")
      {
        alert('Please enter the details')
      }

      //Paid//Unpaid//Voluntary

      else if(this.state.MyDornamantKey=="")
      {

        alert('Please select active or dornamant Company')
      }

      else if(this.state.MyFinancialService=="")
      {
        alert('Please select Financial Service')

      }

      else if(this.state.MyBoardofDirectorsYesorNo=="")
      {

        alert('Please select do you want to sit as Board of director or not')
      }

      else if(this.state.StockValue=="")
      {

      alert('Please enter the stock value')
      }

      else if(this.state.insideryesorNo=='')
      {

        alert('Please select Insider trader Yes or No')
      }

     
      if(this.state.insideryesorNo=='Yes')
      {
  
        if(this.state.yourFullName=='')
        {
  
       alert('Please enter your FullName')
        }
  
        else if(this.state.dtinsider=="")
        {
  
          alert('Please select date of Insider')
        }
  
        else if(this.state.clientname==="")
        {
          alert('Please enter the client name')
        }
  
        else if(this.state.clientproject=="")
        {
  
          alert('Please enter the clientproject')
        }
  
        else if(this.state.whomesclated=="")
        {
  
          alert('Please enter whom it is esclated')
        }
  
        else if(this.state.dtpublic=="")
        {
  
          alert('Please select date of public')
  
        }

        if(this.state.MyConflictYesNo=='')
        {

            alert('Please select other conflict of Intrest')
        }

        if(this.state.MyConflictYesNo=='Yes')
        {
    
          if(this.state.detailsconflict=='')
          {
    
          alert('Please enter the details of Conflit')
          }

          else
          {

            let FinaldtInsider='';

            let FinaldtPublic='';

            if(this.state.dtinsider!="")
            {
      
                  let date1=(this.state.dtinsider.getDate()+1);
      
                  console.log(date1);
                  
                  let month1= (this.state.dtinsider.getMonth()+1);
                  
                  let year1 =(this.state.dtinsider.getFullYear());
                  
                  FinaldtInsider=month1+'/'+this.state.dtinsider.getDate() +'/' +year1;
      
                  console.log(FinaldtInsider);
      
            }
      
            if(this.state.dtpublic!="")
            {
      
                 let date2=(this.state.dtpublic.getDate()+1);
      
                  console.log(date2);
                  
                  let month2= (this.state.dtpublic.getMonth()+1);
                  
                  let year2 =(this.state.dtpublic.getFullYear());
                  
                  FinaldtPublic=month2+'/'+this.state.dtpublic.getDate() +'/' +year2;
                  
      
            }

           
            this._service.Save(
              
            this.state.Contactname,
            this.state.MyBussinessUnitValue,

            //WorkEmployment
           
            this.state.MyEmploymentYesNo,
            this.state.Companyname,
            this.state.RegisterNumber,
            this.state.Position,
            this.state.details,
            paidorunpaidCheckboxvalue,
            this.state.MyDornamantValue,
            this.state.MyFinancialServiceText,
            this.state.MyBoardofDirectorsYesorNo,
             BoardofDirectorsCheckboxValue,
            //End
            //ShareOwnership
            this.state.StockValue,
            //End

            //InsiderInformation

            this.state.insideryesorNo,
            this.state.yourFullName,
            FinaldtInsider,
            this.state.clientname,
            this.state.clientproject,
            this.state.whomesclated,
            FinaldtPublic,
            //End

            //Conflict
            this.state.MyConflictYesNo,
            this.state.detailsconflict,

            //End

            //LastValues

            BussinessValuetext

            //End


              
              ).then(function (data:any)
              {
          
                console.log(data);
    
               
                
    
                alert('Record submitted successfully');
 
                window.location.replace(Envval);
                
              });
                  
          }

        }

        else if(this.state.MyConflictYesNo=='No')
        {

        
          
            let FinaldtInsider='';

            let FinaldtPublic='';

            if(this.state.dtinsider!="")
            {
      
                  let date1=(this.state.dtinsider.getDate()+1);
      
                  console.log(date1);
                  
                  let month1= (this.state.dtinsider.getMonth()+1);
                  
                  let year1 =(this.state.dtinsider.getFullYear());
                  
                  FinaldtInsider=month1+'/'+this.state.dtinsider.getDate() +'/' +year1;
      
                  console.log(FinaldtInsider);
      
            }
      
            if(this.state.dtpublic!="")
            {
      
                 let date2=(this.state.dtpublic.getDate()+1);
      
                  console.log(date2);
                  
                  let month2= (this.state.dtpublic.getMonth()+1);
                  
                  let year2 =(this.state.dtpublic.getFullYear());
                  
                  FinaldtPublic=month2+'/'+this.state.dtpublic.getDate() +'/' +year2;
                  
      
            }

           
            this._service.Save(
              
            this.state.Contactname,
            this.state.MyBussinessUnitValue,

            //WorkEmployment
           
            this.state.MyEmploymentYesNo,
            this.state.Companyname,
            this.state.RegisterNumber,
            this.state.Position,
            this.state.details,
            paidorunpaidCheckboxvalue,
            this.state.MyDornamantValue,
            this.state.MyFinancialServiceText,
            this.state.MyBoardofDirectorsYesorNo,
             BoardofDirectorsCheckboxValue,
            //End
            //ShareOwnership
            this.state.StockValue,
            //End

            //InsiderInformation

            this.state.insideryesorNo,
            this.state.yourFullName,
            FinaldtInsider,
            this.state.clientname,
            this.state.clientproject,
            this.state.whomesclated,
            FinaldtPublic,
            //End

            //Conflict
            this.state.MyConflictYesNo,
            this.state.detailsconflict,

            //End

            BussinessValuetext
              
              ).then(function (data:any)
              {
          
                console.log(data);

                
    
               alert('Record submitted successfully ');

               window.location.replace(Envval);
                
              });
                  
          }


        }
 
       
      

    else if(this.state.insideryesorNo=='No')
    {

      if(this.state.MyConflictYesNo=='')
        {

        alert('Please select other Conflit of Intrest')
          
        }


  else if(this.state.MyConflictYesNo=='Yes')
    {

      if(this.state.detailsconflict=='')
      {

      alert('Please enter the details of Conflit')

      }

      else
      {

        let FinaldtInsider;

        let FinaldtPublic;
  
        if(this.state.dtinsider!="")
        {
  
              let date1=(this.state.dtinsider.getDate()+1);
  
              console.log(date1);
              
              let month1= (this.state.dtinsider.getMonth()+1);
              
              let year1 =(this.state.dtinsider.getFullYear());
              
              FinaldtInsider=month1+'/'+this.state.dtinsider.getDate() +'/' +year1;
  
              console.log(FinaldtInsider);
  
        }
  
        if(this.state.dtpublic!="")
        {
  
             let date2=(this.state.dtpublic.getDate()+1);
  
              console.log(date2);
              
              let month2= (this.state.dtpublic.getMonth()+1);
              
              let year2 =(this.state.dtpublic.getFullYear());
              
              FinaldtPublic=month2+'/'+this.state.dtpublic.getDate() +'/' +year2;
              
  
        }
          

        //2nd Copy1

        
        this._service.Save(
              
          this.state.Contactname,
          this.state.MyBussinessUnitValue,

          //WorkEmployment
         
          this.state.MyEmploymentYesNo,
          this.state.Companyname,
          this.state.RegisterNumber,
          this.state.Position,
          this.state.details,
          paidorunpaidCheckboxvalue,
          this.state.MyDornamantValue,
          this.state.MyFinancialServiceText,
          this.state.MyBoardofDirectorsYesorNo,
           BoardofDirectorsCheckboxValue,
          //End
          //ShareOwnership
          this.state.StockValue,
          //End

          //InsiderInformation

          this.state.insideryesorNo,
          this.state.yourFullName,
          FinaldtInsider,
          this.state.clientname,
          this.state.clientproject,
          this.state.whomesclated,
          FinaldtPublic,
          //End

          //Conflict
          this.state.MyConflictYesNo,
          this.state.detailsconflict,

          //End

          BussinessValuetext
            
            ).then(function (data:any)
            {
        
              console.log(data);

              
             alert('Record submitted successfully ');

             window.location.replace(Envval);
              
            });
                


      }



    }

    else if(this.state.MyConflictYesNo=='No')
    {

      let FinaldtInsider;

      let FinaldtPublic;

      if(this.state.dtinsider!="")
      {

            let date1=(this.state.dtinsider.getDate()+1);

            console.log(date1);
            
            let month1= (this.state.dtinsider.getMonth()+1);
            
            let year1 =(this.state.dtinsider.getFullYear());
            
            FinaldtInsider=month1+'/'+this.state.dtinsider.getDate() +'/' +year1;

            console.log(FinaldtInsider);

      }

      if(this.state.dtpublic!="")
      {

           let date2=(this.state.dtpublic.getDate()+1);

            console.log(date2);
            
            let month2= (this.state.dtpublic.getMonth()+1);
            
            let year2 =(this.state.dtpublic.getFullYear());
            
            FinaldtPublic=month2+'/'+this.state.dtpublic.getDate() +'/' +year2;
            

      }
        
           
            this._service.Save(
              
            this.state.Contactname,
            this.state.MyBussinessUnitValue,

            //WorkEmployment
           
            this.state.MyEmploymentYesNo,
            this.state.Companyname,
            this.state.RegisterNumber,
            this.state.Position,
            this.state.details,
            //PaidCheckbox
            //Close
            this.state.MyDornamantValue,
            this.state.MyFinancialServiceText,
            this.state.MyBoardofDirectorsYesorNo,
            //BoardofDirectoryCheckbox
            //Close
            //End

            //ShareOwnership
            this.state.StockValue,
            //End

            //InsiderInformation

            this.state.insideryesorNo,
            this.state.yourFullName,
            FinaldtInsider,
            this.state.clientname,
            this.state.clientproject,
            this.state.whomesclated,
            FinaldtPublic,
            //End

            //Conflict
            this.state.MyConflictYesNo,
            this.state.detailsconflict,

            //End

            BussinessValuetext
              
              ).then(function (data:any)
              {
          
                console.log(data);

                
    
               alert('Record submitted successfully');

               window.location.replace(Envval);
                
              });
      
      



    }



    }


    }

   else if(this.state.MyEmploymentYesNo=='No')
    {

      if(this.state.StockValue=="")
      {

      alert('Please enter the stock value')
      }

      else if(this.state.insideryesorNo=='')
      {

        alert('Please select Insider trader Yes or No')
      }


      if(this.state.insideryesorNo=='Yes')
      {
  
        if(this.state.yourFullName=='')
        {
  
       alert('Please enter your FullName')
        }
  
        else if(this.state.dtinsider=="")
        {
  
          alert('Please select date of Insider')
        }
  
        else if(this.state.clientname==="")
        {
          alert('Please enter the client name')
        }
  
        else if(this.state.clientproject=="")
        {
  
          alert('Please enter the clientproject')
        }
  
        else if(this.state.whomesclated=="")
        {
  
          alert('Please enter whom it is esclated')
        }
  
        else if(this.state.dtpublic=="")
        {
  
          alert('Please select date of public')
  
        }

        else if(this.state.MyConflictYesNo=="")
        {

          alert('Please select the Conflict value ')

        }

        if(this.state.MyConflictYesNo=="Yes")
        {

          if(this.state.detailsconflict=="")
          {

            alert('please enter details of con1flict')
          }

      else
      {

        let FinaldtInsider;

        let FinaldtPublic;
  
        if(this.state.dtinsider!="")
        {
  
              let date1=(this.state.dtinsider.getDate()+1);
  
              console.log(date1);
              
              let month1= (this.state.dtinsider.getMonth()+1);
              
              let year1 =(this.state.dtinsider.getFullYear());
              
              FinaldtInsider=month1+'/'+this.state.dtinsider.getDate() +'/' +year1;
  
              console.log(FinaldtInsider);
  
        }
  
        if(this.state.dtpublic!="")
        {
  
             let date2=(this.state.dtpublic.getDate()+1);
  
              console.log(date2);
              
              let month2= (this.state.dtpublic.getMonth()+1);
              
              let year2 =(this.state.dtpublic.getFullYear());
              
              FinaldtPublic=month2+'/'+this.state.dtpublic.getDate() +'/' +year2;
              
  
        }
          
             
  
        this._service.Save(
              
          this.state.Contactname,
          this.state.MyBussinessUnitValue,

          //WorkEmployment
         
          this.state.MyEmploymentYesNo,
          this.state.Companyname,
          this.state.RegisterNumber,
          this.state.Position,
          this.state.details,
          paidorunpaidCheckboxvalue,
          this.state.MyDornamantValue,
          this.state.MyFinancialServiceText,
          this.state.MyBoardofDirectorsYesorNo,
           BoardofDirectorsCheckboxValue,
          //End
          //ShareOwnership
          this.state.StockValue,
          //End

          //InsiderInformation

          this.state.insideryesorNo,
          this.state.yourFullName,
          FinaldtInsider,
          this.state.clientname,
          this.state.clientproject,
          this.state.whomesclated,
          FinaldtPublic,
          //End

          //Conflict
          this.state.MyConflictYesNo,
          this.state.detailsconflict,

          //End

          BussinessValuetext
            
            ).then(function (data:any)
            {
        
              console.log(data);
  
             
    
               alert('Record submitted successfully ');

               window.location.replace(Envval);
              
            });

        
      }

          
        }

        else if(this.state.MyConflictYesNo=="No")
        {

          let FinaldtInsider;

        let FinaldtPublic;
  
        if(this.state.dtinsider!="")
        {
  
              let date1=(this.state.dtinsider.getDate()+1);
  
              console.log(date1);
              
              let month1= (this.state.dtinsider.getMonth()+1);
              
              let year1 =(this.state.dtinsider.getFullYear());
              
              FinaldtInsider=month1+'/'+this.state.dtinsider.getDate() +'/' +year1;
  
              console.log(FinaldtInsider);
  
        }
  
        if(this.state.dtpublic!="")
        {
  
             let date2=(this.state.dtpublic.getDate()+1);
  
              console.log(date2);
              
              let month2= (this.state.dtpublic.getMonth()+1);
              
              let year2 =(this.state.dtpublic.getFullYear());
              
              FinaldtPublic=month2+'/'+this.state.dtpublic.getDate() +'/' +year2;
              
  
        }
          
             
  
        this._service.Save(
              
          this.state.Contactname,
          this.state.MyBussinessUnitValue,

          //WorkEmployment
         
          this.state.MyEmploymentYesNo,
          this.state.Companyname,
          this.state.RegisterNumber,
          this.state.Position,
          this.state.details,
          paidorunpaidCheckboxvalue,
          this.state.MyDornamantValue,
          this.state.MyFinancialServiceText,
          this.state.MyBoardofDirectorsYesorNo,
           BoardofDirectorsCheckboxValue,
          //End
          //ShareOwnership
          this.state.StockValue,
          //End

          //InsiderInformation

          this.state.insideryesorNo,
          this.state.yourFullName,
          FinaldtInsider,
          this.state.clientname,
          this.state.clientproject,
          this.state.whomesclated,
          FinaldtPublic,
          //End

          //Conflict
          this.state.MyConflictYesNo,
          this.state.detailsconflict,

          //End
            
          BussinessValuetext

            
            ).then(function (data:any)
            {
        
              console.log(data);
  
              
    
               alert('Record submitted successfully ');

               window.location.replace(Envval);
              
            });
                
        
        }
  
       
      }

    else if(this.state.insideryesorNo=='No')
    {

      if(this.state.MyConflictYesNo=="")
      {

        alert('Please select the Conflict value')

      }

      else if(this.state.MyConflictYesNo=="Yes")
      {

        if(this.state.detailsconflict=="")
        {

          alert('please enter details of con1flict')
        }

        else
        {

          let FinaldtInsider;

          let FinaldtPublic;
    
          if(this.state.dtinsider!="")
          {
    
                let date1=(this.state.dtinsider.getDate()+1);
    
                console.log(date1);
                
                let month1= (this.state.dtinsider.getMonth()+1);
                
                let year1 =(this.state.dtinsider.getFullYear());
                
                FinaldtInsider=month1+'/'+this.state.dtinsider.getDate() +'/' +year1;
    
                console.log(FinaldtInsider);
    
          }
    
          if(this.state.dtpublic!="")
          {
    
               let date2=(this.state.dtpublic.getDate()+1);
    
                console.log(date2);
                
                let month2= (this.state.dtpublic.getMonth()+1);
                
                let year2 =(this.state.dtpublic.getFullYear());
                
                FinaldtPublic=month2+'/'+this.state.dtpublic.getDate() +'/' +year2;
                
    
          }


          //Last Copy

          this._service.Save(
              
            this.state.Contactname,
            this.state.MyBussinessUnitValue,

            //WorkEmployment
           
            this.state.MyEmploymentYesNo,
            this.state.Companyname,
            this.state.RegisterNumber,
            this.state.Position,
            this.state.details,
            paidorunpaidCheckboxvalue,
            this.state.MyDornamantValue,
            this.state.MyFinancialServiceText,
            this.state.MyBoardofDirectorsYesorNo,
             BoardofDirectorsCheckboxValue,
            //End
            //ShareOwnership
            this.state.StockValue,
            //End

            //InsiderInformation

            this.state.insideryesorNo,
            this.state.yourFullName,
            FinaldtInsider,
            this.state.clientname,
            this.state.clientproject,
            this.state.whomesclated,
            FinaldtPublic,
            //End

            //Conflict
            this.state.MyConflictYesNo,
            this.state.detailsconflict,

            //End

            BussinessValuetext
              
              ).then(function (data:any)
              {
          
                console.log(data);
    
                
    
               alert('Record submitted successfully ');

               window.location.replace(Envval);
                
              });
                  

         
        }

      }

      else
      {

        //Write Here

        this._service.SaveAllNOS(this.state.Contactname,
        this.state.MyBussinessUnitValue,
        //WorkEmployment
         this.state.MyEmploymentYesNo,
         //End
         //ShareOwnership

         this.state.StockValue,
         //End

         this.state.insideryesorNo,
         //End

         //Conflict
        this.state.MyConflictYesNo,
        
        //End

        BussinessValuetext

              
          ).then(function (data:any)
          {
      
            console.log(data);

            
    
            alert('Record submitted successfully ');

            window.location.replace(Envval);
            
          });
      }


    }

    }

    //End WorkEmployment

   }

  
  
  public render(): React.ReactElement<IConflictRegistrySubmitProps> {

    return (

<React.Fragment>

{/* <Collapsible trigger="" className={styles.Collapsible__trigger} > */}

<b><label className={styles.Mainheading}>Please provide your contact information</label></b><br/>

<div id="divregion2"> 

<div className={styles.Divsection}>

<b><label className={styles.HeadLable}>Contact Information</label></b><br/>

</div>

<Stack horizontal tokens={stackTokens1}>
<StackItem className={styles.coststyle} >
<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Name <label className={styles.recolorss}>*</label></label></b><br/><br/> 
<input type="text" name="txtName" className={styles.boxsizeContactInfo} value={this.state.Contactname} onChange={this.changeContactName.bind(this)}/><br></br>
</div>
</StackItem>

</Stack>


<StackItem className={styles.coststyle} >
<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Business Unit <label className={styles.recolorss}>*</label></label></b><br/><br/> 
<Dropdown className={styles.onlyFont}
  placeholder="Select  Business Unit"
  options={this.state.BussinessUnitItems}
  
  styles={dropdownStyles}
  selectedKey={this.state.MyBussinessUnitValue ? this.state.MyBussinessUnitValue : undefined} onChange={this.hadleBussinessUnit.bind(this)}/><br></br>
</div>
</StackItem>

</div>
{/* </Collapsible> */}

<br></br>
{/* <Collapsible trigger="Please advise of any outside of work employment/ activities outlined within the acknowledgement above that you are involved in. The noted activities " className={styles.HeadLable}> */}

<div className={styles.Divsection}>

<b><label className={styles.Mainheading}>Outside Work Employment / Activities</label></b><br/>

</div>

<Stack horizontal tokens={stackTokens1}>
<StackItem className={styles.coststyle} >
<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Do you wish to register any other employment or activities that you are involved in outside of your Capco employment? <label className={styles.recolorss}>*</label></label></b><br/><br/> 
<Dropdown className={styles.onlyFont}
  placeholder="Select  Yes or NO"
  options={drpYesorNo}
  styles={dropdownStyles}
  selectedKey={this.state.MyEmploymentYesNo ? this.state.MyEmploymentYesNo : undefined} onChange={this.handleEmploymentYesNo.bind(this)}/><br></br>
</div>
</StackItem>

</Stack>

{this.state.MyWorkEmploymentboolean == true  &&

<Stack horizontal tokens={stackTokens1}>
<StackItem className={styles.coststyle} >
<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Company Name <label className={styles.recolorss}>*</label></label></b><br/><br/> 
<input type="text" name="txtCompanyName" className={styles.boxsize} value={this.state.Companyname} onChange={this.changeCompanyName.bind(this)}/><br></br>
</div><br></br>

<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}> Type of Company and Registration Number<label className={styles.recolorss}>*</label></label></b><br/><br/> 
<input type="text" name="txttypecmpRegNumber" className={styles.boxsize} value={this.state.RegisterNumber} onChange={this.changeRegisterNumber.bind(this)}/><br></br>
</div>

<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}> Position<label className={styles.recolorss}>*</label></label></b><br/><br/> 
<input type="text" name="txtPosition" className={styles.boxsize} value={this.state.Position} onChange={this.changePosition.bind(this)}/><br></br>
</div>

<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Details<label className={styles.recolorss}>*</label></label></b><br/><br/> 
<textarea id="txtDetails" value={this.state.details} onChange={this.changedetails.bind(this)} className={styles.boxsizemultiline}></textarea>
</div>

<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Paid / Unpaid / Voluntary<label className={styles.recolorss}>*</label></label></b><br/><br/> 
{this.state.choiceValues.map((choiceValue: string) => {


                  return (
                    <div style={{ margin: "2px", padding: "3px" }}>
                      <Checkbox style={{ margin: "2px", padding: "3px" }} label={choiceValue} onChange={this._onCheckboxEmployment} />
                    </div>
                  );
                }
                )}
</div>

<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Is this an active or dormant company?<label className={styles.recolorss}>*</label></label></b><br/><br/> 
<Dropdown className={styles.onlyFont}
  placeholder="Select  active or dormant company"
  options={this.state.CompanyTypes}
  styles={dropdownStyles1}
  selectedKey={this.state.MyDornamantKey ? this.state.MyDornamantKey : undefined} onChange={this.handleActiveorDornamant.bind(this)}/><br></br>
</div>

<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Is the company involved in financial services included but not limited to: banking, hedge funds, real estate, private equity or financial technology?<label className={styles.recolorss}>*</label></label></b><br/><br/> 
<Dropdown className={styles.onlyFont}
  placeholder="Select company involved financial services"
  options={this.state.ComanyFinancialServiceTypes}
  styles={dropdownStyles1}
  selectedKey={this.state.MyFinancialService ? this.state.MyFinancialService : undefined} onChange={this.handlefinancialservice.bind(this)}/><br></br>
</div>

<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Do you sit on the Board of Directors?<label className={styles.recolorss}>*</label></label></b><br/><br/> 
<Dropdown className={styles.onlyFont}
  placeholder="Select Yes or No"
  options={drpYesorNo}
  styles={dropdownStyles1}
  selectedKey={this.state.MyBoardofDirectorsYesorNo ? this.state.MyBoardofDirectorsYesorNo : undefined} onChange={this.handleBoardofDirectorsYesNo.bind(this)}/><br></br>
</div>


{this.state.Myboardofdirecrosboolean == true  &&

<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>As a Board of Director, are you on a Board with a Capco client or an employee of a Capco client?<label className={styles.recolorss}>*</label></label></b><br/><br/> 
{this.state.choiceValues1.map((choiceValueboardofdirectors: string) => {


                  return (
                    <div style={{ margin: "2px", padding: "3px" }}>
                      <Checkbox style={{ margin: "2px", padding: "3px" }} label={choiceValueboardofdirectors} onChange={this._onCheckboxBoradofDirectors} />
                    </div>
                  );
                }
                )}
</div>
  
}

</StackItem>
</Stack>

}



{/* </Collapsible> */}

<br></br>

{/* <Collapsible trigger="Share Ownership" className={styles.HeadLable}> */}

<b><label className={styles.Mainheading}>Share Ownership</label></b><br/>
<div className={styles.Divsection}>

<b><label className={styles.labelsFonts}>Do you maintain ownership of a company stock greater than 5%?</label></b><br/><br/> 
<input type="text" name="txtstockvalue" className={styles.boxsize} value={this.state.StockValue} onChange={this.changeStockvalue.bind(this)}/><br></br>
</div>

{/* </Collapsible> */}

<br></br>

{/* <Collapsible trigger="Insider List Description" className={styles.HeadLable}> */}

<b><label className={styles.Mainheading}> Insider Trading List Identification</label></b><br/>


<div className={styles.Divsection}>

<b><label className={styles.HeadLable}>Please review the below definition regarding insider trading list</label></b><br/>

</div>

{/* <div className={styles.Divsection}>

<b><label className={styles.HeadLable}>Insider List</label></b><br/>

</div> */}


<p className={styles.HeadPara}>As required by the Market Abuse Regulation (Regulation 596/2014); a list maintained by issuers (or any person<br></br>   
acting on their behalf or on their account) which includes details of all persons who have access to inside<br></br>     
information and who are working for them under a contract of employment, or otherwise performing tasks<br></br>
through which they have access to inside information.

</p>

<p className={styles.HeadPara}>For the purposes of the Market Abuse Regulation (Regulation 596/2014), information of a precise nature, that:</p>

<ul>
  <li className={styles.HeadPara}>Has not been made public</li>
  <li className={styles.HeadPara}>Relates, directly or indirectly, to one or more issuers or to one or more financial instruments; and</li>
  <li className={styles.HeadPara}>If it were made public, would be likely to have a significant effect on the prices of those financial instruments or on the price of related derivative financial instruments (that is, it is information that a reasonable investor would be likely to use as part of the basis of their investment decisions).</li>
</ul>  



{/* </Collapsible> */}

<br></br>

{/* <b><label className={styles.HeadLable}>Insider Trading List Identification</label></b><br/> */}


{/* <div className={styles.Divsection}>
<b><label className={styles.HeadLable}>Insider Information</label></b><br></br>
</div> */}

<Stack horizontal tokens={stackTokens1}>
<StackItem className={styles.coststyle} >
<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Have you been added as part of an insider's trader list? <label className={styles.recolorss}>*</label></label></b><br/><br/> 
<Dropdown className={styles.onlyFont}
  placeholder="Select  Yes or NO"
  options={drpYesorNo}
  styles={dropdownStyles}
  selectedKey={this.state.insideryesorNo ? this.state.insideryesorNo : undefined} onChange={this.handleInsiderYesorNo.bind(this)}/><br></br>
</div>

{this.state.insideryesorNoboolean==true &&

<Stack horizontal tokens={stackTokens1}>
<StackItem className={styles.coststyle} >
<div className={styles.Divsection}>
<b><label className={styles.HeadLable}>Insider Information</label></b><br/>
</div>

<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Your Full Name <label className={styles.recolorss}>*</label></label></b><br/><br/> 
<input type="text" name="txtfullName" className={styles.boxsize} value={this.state.yourFullName} onChange={this.changeFullName.bind(this)}/><br></br>
<br></br>


<b><label className={styles.HeadLable}>Please select the date you became an insider.</label></b><br></br><br></br>

<DatePicker placeholder="Select a date..."
                            onSelectDate={this.handldtInsiderDateChange}
                            value={this.state.dtinsider}
                            formatDate={this._onFormatDate1}
                            isMonthPickerVisible={false}
                            className={styles.boxsize}
                            />
  
                            <br></br>

<b><label className={styles.labelsFonts}>Client Name <label className={styles.recolorss}>*</label></label></b><br/><br/> 
<input type="text" name="txtClientname" className={styles.boxsize} value={this.state.clientname} onChange={this.changeClientname.bind(this)}/><br></br>
<br></br>

<b><label className={styles.labelsFonts}>Client project name or code phrase <label className={styles.recolorss}>*</label></label></b><br/><br/> 
<input type="text" name="txtprojectname" className={styles.boxsize} value={this.state.clientproject} onChange={this.changeProjectName.bind(this)}/><br></br>
<br></br>

<b><label className={styles.labelsFonts}>To whom was this escalated? <label className={styles.recolorss}>*</label></label></b><br/><br/> 
<input type="text" name="txtescalated" className={styles.boxsize} value={this.state.whomesclated} onChange={this.changewhotoesclated.bind(this)}/><br></br>
<br></br>

<b><label className={styles.HeadLable}> Please provide the date the insider’s list will cease to exist or anticipated date the information will become public.</label></b><br></br><br></br>

<DatePicker placeholder="Select a date..."
                            onSelectDate={this.handldtPublicDateChange}
                            value={this.state.dtpublic}
                            formatDate={this._onFormatDate1}
                            isMonthPickerVisible={false}
                            className={styles.boxsize}
                            />
</div>

                            <br></br>

</StackItem>
</Stack>
}



</StackItem>

</Stack>



{/* </Collapsible> */}

<br></br>

<b><label className={styles.Mainheading}>Other Conflicts</label></b><br/>

{/* <Collapsible trigger="Other Conflicts" className={styles.HeadLable}> */}

<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}> Are there any other conflicts of interest?<label className={styles.recolorss}>*</label></label></b><br/><br/> 
<Dropdown className={styles.onlyFont}
  placeholder="Select Yes or No"
  options={drpYesorNo}
  styles={dropdownStyles1}
  selectedKey={this.state.MyConflictYesNo ? this.state.MyConflictYesNo : undefined} onChange={this.handleConflictsYesNo.bind(this)}/><br></br>
</div>
{this.state.MyConflictbollean==true &&

<div className={styles.Divsection}> 
<b><label className={styles.labelsFonts}>Please provide details of conflict<label className={styles.recolorss}>*</label></label></b><br/><br/> 
<textarea id="txtdetailsconflict" value={this.state.detailsconflict} onChange={this.changedetailsconflict.bind(this)} className={styles.boxsizemultiline}></textarea>
</div>


}
<div className={styles.Divsection}> 
<PrimaryButton text="Submit" onClick={this.OnBtnClickSubmit.bind(this)} styles={stackButtonStyles} className={styles.Approvebutton}/><br></br>
</div>

{/* </Collapsible> */}

</React.Fragment>

);





  }
}
