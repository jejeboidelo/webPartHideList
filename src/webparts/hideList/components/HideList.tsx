import * as React from 'react';
import styles from './HideList.module.scss';
import { IHideListProps } from './IHideListProps';
import { getSP } from '../pnpjsConfig';
import { SPFI } from '@pnp/sp';
import "react-table-6/react-table.css" 
import { PermissionKind } from '@pnp/sp/security';
import { Provider, Table, teamsTheme, Button, Pill} from '@fluentui/react-northstar';


const header = {
  key: 'header',
  items: [
      {
          content: 'Nom de la liste',
          key: 'nom'
      },
      {
          content: 'Id de la liste',
          key: 'id'
      },
      {
          content: 'Description de la liste',
          key: 'description'
      },
      {
          content: 'Etat',
          key: 'etat'
      },
      {
          content: "Action",
          key: 'action'
      }
  ]
  
}

interface IHideListState {
  isAdmin:boolean,
  liste: any[]
}

export default class HideList extends React.Component<IHideListProps, IHideListState> {

  private _sp: SPFI;

  constructor(props: IHideListProps){
    super(props);
    this.state = {
      isAdmin: false,
      liste: []
    };
    
    
    this._sp = getSP();

    this.getList = this.getList.bind(this);
    this.isAdmin = this.isAdmin.bind(this);
    this.hideList = this.hideList.bind(this);
    this.unhideList = this.unhideList.bind(this);
  
    
  }

  public componentDidMount(): void {
    this.getList();
    this.isAdmin();
  }



  public render(): React.ReactElement<IHideListProps> {
    const {
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.hideList} ${hasTeamsContext ? styles.teams : ''}`}>
        <Provider theme={teamsTheme}>
          <Table header={header} rows={this.state.liste}/>
        </Provider>
      </section>
    );
  }

  public async isAdmin() {
    const res = this._sp.web.currentUserHasPermissions(PermissionKind.AddDelPrivateWebParts);

    res.then(res => this.setState({isAdmin: res}));
  }


  public getList(): any { 
    const res = this._sp.web.lists();
    var temp:any[] = [];
    
    res.then((result) => {
      result.forEach((list, i) => {

        if(this.props.BaseType.some((element)=>{return element == list.BaseType;}) &&
          this.props.ListeTemplate.some((element)=>{return Number(element) == list.BaseTemplate || element == "-9999"})){
            temp.push(
              {
                key:i+1,
                items:[
                  {content: list.Title,
                    truncateContent: true},
                  { content: list.Id,
                    truncateContent: true},
                  {content: list.Description,
                    truncateContent: true},
                  {content: <Pill style={!list.Hidden ? {backgroundColor: "#87ceeb"}: {backgroundColor: "#ff6347"}}>
                    {list.Hidden ? "Cachée":"Visible"}</Pill>},
                  {
                    content: 
                      list.Hidden ? 
                          <Button disabled={!this.state.isAdmin} onClick={()=>this.unhideList(list.Id)}>Afficher</Button>:
                          <Button disabled={!this.state.isAdmin} onClick={()=>this.hideList(list.Id)}>Cacher</Button>}
                ]
              }
            );
          }
      });

      this.setState({
        liste: temp
      })
    })
  }

  public async hideList(id: string){
    try {
      await this._sp.web.lists.getById(id).update({Hidden: true})
      .catch((err:any)=>{
        if(err.status==403){
          alert("Impossible de cacher cette liste")}
        else{
          alert("Une erreur est survenue")}
        })
      this.getList();
    }
    catch (e){
      console.log(e);
    }
  }

  public async unhideList(id: string){
    try {
      await this._sp.web.lists.getById(id).update({Hidden: false})
        .catch((err:any)=>{
          if(err.status==403){
            alert("Impossible d'afficher cette liste")}
          else{
            alert("Une erreur est survenue")}
          })
      this.getList();
    }
    catch (e){
      console.log(e);
    }
  }

  componentDidUpdate(prevProps: Readonly<IHideListProps>, prevState: Readonly<IHideListState>, snapshot?: any): void {
    if(this.props.BaseType!=prevProps.BaseType || this.props.ListeTemplate!=prevProps.ListeTemplate){    
      this.getList();
    }
  }
}