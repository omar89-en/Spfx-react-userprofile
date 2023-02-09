import * as React from "react";
import styles from "./UserProfileViewer.module.scss";
import { IUserProfileViewerProps } from "./IUserProfileViewerProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { IUserProfile } from "./IUserProfile";
import { IUserProfileViewerState } from "./IUserProfileViewerState";
import { IDataService } from "../services/IDataService";
import { ServiceScope } from "@microsoft/sp-core-library";
import { UserProfileService } from "../services/UserProfileService";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faEnvelope } from "@fortawesome/free-regular-svg-icons";
import { faFolder } from "@fortawesome/free-regular-svg-icons";
import { FontIcon, Icon } from "office-ui-fabric-react/lib/Icon";
import {
  Persona,
  PersonaSize,
  PersonaPresence
} from 'office-ui-fabric-react/lib/Persona';
import { LivePersonaCard } from '../../../controls/LivePersonaCard';
export class UserProfile implements IUserProfile {
  FirstName: string;
  LastName: string;
  Email: string;
  Title: string;
  WorkPhone: string;
  DisplayName: string;
  Department: string;
  HireDate: string;
  PictureURL: string;
  Office:string;
  PersonalSiteInstantiationState: string;

  UserProfileProperties: Array<any>;
}

export default class UserProfileViewer extends React.Component<
  IUserProfileViewerProps,
  IUserProfileViewerState
> {
  private dataCenterServiceInstance: IDataService;

  constructor(props: IUserProfileViewerProps, state: IUserProfileViewerState) {
    super(props);

    let userProfile: IUserProfile = new UserProfile();
    userProfile.FirstName = "";
    userProfile.LastName = "";
    userProfile.Email = "";
    userProfile.Title = "";
    userProfile.WorkPhone = "";
    userProfile.DisplayName = "";
    userProfile.Department = "";
    userProfile.PictureURL = "";
    userProfile.HireDate = "";
    userProfile.PersonalSiteInstantiationState = "";
    userProfile.Office=""

    userProfile.UserProfileProperties = [];

    this.state = {
      userProfileItems: userProfile,
    };
  }

  public componentWillMount(): void {
    let serviceScope: ServiceScope = this.props.serviceScope;
    this.dataCenterServiceInstance = serviceScope.consume(
      UserProfileService.serviceKey
    );

    this.dataCenterServiceInstance
      .getUserProfileProperties()
      .then((userProfileItems: IUserProfile) => {
        for (
          let i: number = 0;
          i < userProfileItems.UserProfileProperties.length;
          i++
        ) {
          if (userProfileItems.UserProfileProperties[i].Key == "FirstName") {
            userProfileItems.FirstName =
              userProfileItems.UserProfileProperties[i].Value;
          }

          if (userProfileItems.UserProfileProperties[i].Key == "LastName") {
            userProfileItems.LastName =
              userProfileItems.UserProfileProperties[i].Value;
          }

          if (userProfileItems.UserProfileProperties[i].Key == "WorkPhone") {
            userProfileItems.WorkPhone =
              userProfileItems.UserProfileProperties[i].Value;
          }

          if (userProfileItems.UserProfileProperties[i].Key == "Department") {
            userProfileItems.Department =
              userProfileItems.UserProfileProperties[i].Value;
          }

          if (userProfileItems.UserProfileProperties[i].Key == "PictureURL") {
            userProfileItems.PictureURL =
              userProfileItems.UserProfileProperties[i].Value;
          }
          if (userProfileItems.UserProfileProperties[i].Key == "Email") {
            userProfileItems.Email =
              userProfileItems.UserProfileProperties[i].Value;
          }

          if (
            userProfileItems.UserProfileProperties[i].Key ==
            "SPS-PersonalSiteInstantiationState"
          ) {
            userProfileItems.PersonalSiteInstantiationState =
              userProfileItems.UserProfileProperties[i].Value;
          }
          if (userProfileItems.UserProfileProperties[i].Key == "DisplayName") {
            userProfileItems.DisplayName =
              userProfileItems.UserProfileProperties[i].Value;
          }

          if (userProfileItems.UserProfileProperties[i].Key == "SPS-HireDate") {
            userProfileItems.HireDate =
              userProfileItems.UserProfileProperties[i].Value;
          }

          if (userProfileItems.UserProfileProperties[i].Key == "Office") {
            userProfileItems.Office =
              userProfileItems.UserProfileProperties[i].Value;
          }
          
        }

        this.setState({ userProfileItems: userProfileItems });
      });
  }
 private getInitials(fullName: string): string {
    if (!fullName) {
      return (null);
    }

    let parts: string[] = fullName.split(' ');

    let initials: string = "";
    parts.forEach(p => {
      if (p.length > 0) {
          initials = initials.concat(p.substring(0, 1).toUpperCase());
      }
    });

    return (initials);
  }

//personne Card avec style Css
/*public render(): React.ReactElement<IUserProfileViewerProps> {

    



  return (
    <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.showname}>
              {this.state.userProfileItems.DisplayName}
            </div>

            <div className={styles.overlay}>
              <div className={styles.card}>
                <div className={styles.chip}>
                  <img
                    src={this.state.userProfileItems.PictureURL}
                    width="60"
                    height="60"
                  ></img>
                  {this.state.userProfileItems.DisplayName}
                  <p className={styles.title}>
                    {this.state.userProfileItems.Department}
                  </p>
                </div>

                <div className={styles.info}>
                  <p className={styles.Contact}>
                    
                    Contact
                    <Icon
                      iconName="ChevronRightSmall"
                      className={styles.ChevronRightSmall}
                    />
                  </p>

                  {this.state.userProfileItems.Email.length > 0 && (
                    <p >
                      <FontAwesomeIcon icon={faEnvelope} className={styles.esp} />
                      {this.state.userProfileItems.Email}
                    </p>
                  )}

                  {this.state.userProfileItems.WorkPhone.length > 0 && (
                    <p>
                      <Icon iconName="Phone" className={styles.esp}  />
                      {this.state.userProfileItems.WorkPhone}
                    </p>
                  )}

                  {this.state.userProfileItems.HireDate.length > 0 && (
                    <p>
                      <Icon iconName="DateTime" className={styles.esp} />
                      {this.state.userProfileItems.HireDate}
                    </p>
                  )}
                   {this.state.userProfileItems.Office.length > 0 && (
                    <p>
                      <Icon iconName="MapPin" className={styles.esp} />
                      {this.state.userProfileItems.Office}
                    </p>
                  )}

                </div>
              </div>
            </div>
          </div>
        </div>
    
  );
}*/
 
//personne Card avec LivePersonaCard
public render(): React.ReactElement<IUserProfileViewerProps> {


  

    return (
      <div>
        
      
      
        <div className={styles.lpcSample}>
       
       <LivePersonaCard
        user={{
          displayName: this.state.userProfileItems.DisplayName,
          email: this.state.userProfileItems.Email,
        }}
       serviceScope={this.props.context.serviceScope}
       

/>
       


      </div>
      </div>
      
    );
  }
}
