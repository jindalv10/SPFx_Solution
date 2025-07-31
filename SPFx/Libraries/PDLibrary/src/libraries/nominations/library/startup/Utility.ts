import { AllRoles } from "../../models/IConstants";
import { ISpUser } from "../../models/ISpUser";
export class Utility {

    public static getRolesForFolder(): string[] {
        let allRolesForAttachments = [];
        Object.keys(AllRoles).map(key => {
            if (AllRoles[key] != AllRoles.LA.toUpperCase())
                allRolesForAttachments.push(AllRoles[key]);
        });
        return allRolesForAttachments;
    }
    public static getDocSetName(nominee: ISpUser, id: number): string {

        let docSetName = null;
        docSetName = nominee && nominee.title + "-" + id;
        return docSetName;
    }

    public static findDeepNestedObject(objectVal, searchKey: string):string {
      searchKey = searchKey.replace(/\[(\w+)\]/g, '.$1');
      searchKey = searchKey.replace(/^\./, '');
      var a = searchKey.split('.');
      for (var i = 0, n = a.length; i < n; ++i) {
          var k = a[i];

          if (k in objectVal) {
            objectVal = objectVal[k];
          } else {
              return;
          }

      }

      if(Array.isArray(objectVal) && objectVal.length > 0){
        return objectVal.map(elem => {return elem.hasOwnProperty('title') ? elem.title : elem;}).join(", ");}
        else{
          return objectVal;
      }
      //return objectVal;
  }

  public static isObjectNullOrEmpty = (obj: any) => {
    return (obj == null || obj == undefined);
  }

  public static mergeByID(firstKey: string, secondKey:string, array1, array2) {
    return array1.map(elemArr1 => {
      let array2Elem = {baseKey: array2.find(element => element[secondKey] ===  parseInt(elemArr1[firstKey]))};
      return array2Elem.baseKey !== undefined ? { ...elemArr1, ...array2Elem } : { ...elemArr1 };
    });
  }



  public static ciEquals(a, b) {
    return typeof a === 'string' && typeof b === 'string'
        ? a.localeCompare(b, undefined, { sensitivity: 'accent' }) === 0
        : a === b;
  }

}

