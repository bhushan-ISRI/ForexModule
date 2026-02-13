// import { IForexModuleProps } from '../../components/IForexModuleProps';
// import SPCRUDOPS from '../DAL/spcrudops';
// import { IUserProfile } from '../../services/INTERFACES/IUserProfile';


// export interface UserProfileOps {
//     getLoggUserProfile(props: IForexModuleProps): Promise<IUserProfile>;
// }

// export default function LoggUserProfileOps() {
//     const spCrudOps = SPCRUDOPS();

//     // const getVendorMaster = async (brrId: string | number, props: IForexModuleProps): Promise<IUserProfile> => {
//         const getLoggUserProfile = async (props: IForexModuleProps): Promise<IUserProfile> => {
//             return await (await spCrudOps).currentProfile(props).then(results => {
//                     let brr: Array<IUserProfile> = new Array<IUserProfile>();
//                     if(results !== undefined){
//                         brr.push({
//                             AccountName: results.AccountName,
//                             UserProfileProperties:results.UserProfileProperties!==undefined?results.UserProfileProperties:[],
//                             Location:results.UserProfileProperties.find(prop => prop.Key === 'Office')?.Value || "Location not found"
//                         });
//                     }
//                     return brr[0];
 
                        
                    
//                 }
//                 );
//         //});
//     };





//     return {
//         getLoggUserProfile
//     };
// }

import { IForexModuleProps } from '../../components/IForexModuleProps';
import SPCRUDOPS from '../DAL/spcrudops';
import { IUserProfile } from '../../service/INTERFACE/IUserProfile';

export interface UserProfileOps {
    getLoggUserProfile(props: IForexModuleProps): Promise<IUserProfile>;
}

export default function LoggUserProfileOps(): UserProfileOps {
    const spCrudOps = SPCRUDOPS();

    const getLoggUserProfile = async (props: IForexModuleProps): Promise<IUserProfile> => {
        const result = await (await spCrudOps).currentProfile(props);

        // Debug log to inspect all properties
        //console.log("UserProfilePropertiesRS:", result.UserProfileProperties);

        const userProfileProperties = result.UserProfileProperties ?? [];

        // Attempt to find the location using a likely key
        const locationValue = userProfileProperties.find(
            ( prop: { Key: string; }) => prop.Key === 'Office' // Change this key if needed
        )?.Value || "Location not found";

        return {
            AccountName: result.AccountName,
            UserProfileProperties: userProfileProperties,
            Location: locationValue
        };
    };

    return {
        getLoggUserProfile
    };
}
