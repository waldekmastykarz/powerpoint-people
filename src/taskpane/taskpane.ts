/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import { IDynamicPerson, Providers, ProviderState, SimpleProvider } from "@microsoft/mgt";
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global $, document, Office */

import { getAccessToken, signIn } from "./../helpers/ssoauthhelper";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    $(document).ready(function () {
      $("#signin").click(signIn);
    });
  }
});

export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];
  let userProfileInfo: string[] = [];
  userProfileInfo.push(result["displayName"]);
  userProfileInfo.push(result["jobTitle"]);
  userProfileInfo.push(result["mail"]);
  userProfileInfo.push(result["mobilePhone"]);
  userProfileInfo.push(result["officeLocation"]);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }
  Office.context.document.setSelectedDataAsync(userInfo, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}

let imageCount = 0;

export function showPeoplePicker(): void {
  Providers.globalProvider = new SimpleProvider(() => {
    return Promise.resolve(getAccessToken());
  }, () => {
    Providers.globalProvider.setState(ProviderState.SignedIn);
    return Promise.resolve();
  }, () => {
    Providers.globalProvider.setState(ProviderState.SignedOut);
    return Promise.resolve();
  });
  Providers.globalProvider.setState(ProviderState.SignedIn);
  $("#content").append('<mgt-people-picker type="Person"></mgt-people-picker>');
  $("#content").append("<button>Insert</button>");
  $("#content button").click(() => {
    const people = [...(($("#content mgt-people-picker")[0] as any).selectedPeople as IDynamicPerson[])];
    imageCount = 0;
    addPeopleInfo(people);
  });
}

function addPeopleInfo(people: IDynamicPerson[]) {
  const person = people.shift();
  addPersonInfo(person).then(
    () => {
      if (people.length > 0) {
        imageCount++;
        addPeopleInfo(people);
      }
    },
    (err) => console.error(err)
  );
}

function addPersonInfo(person: IDynamicPerson): Promise<void> {
  return new Promise<void>((resolve, reject): void => {
    const p = `${person.displayName}\n${person.jobTitle}`;
    Office.context.document.setSelectedDataAsync(p, (res) => {
      if (res.error) {
        return reject(res.error);
      }

      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: Office.ValueFormat.Formatted },
        (res) => {
          if (res.error) {
            return reject(res.error);
          }

          Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              return reject("Action failed with error: " + asyncResult.error.message);
            }

            const currentSlideId = (asyncResult.value as any).slides[0].id;
            Office.context.document.goToByIdAsync(currentSlideId, Office.GoToType.Slide, function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                return reject("Action failed with error: " + asyncResult.error.message);
              }

              if (!person.personImage) {
                return resolve();
              }

              Office.context.document.setSelectedDataAsync(
                person.personImage.substr(person.personImage.indexOf(",") + 1),
                {
                  coercionType: Office.CoercionType.Image,
                  imageWidth: 100,
                  imageLeft: 150 + 1000 * imageCount,
                  imageTop: 400,
                },
                (res) => {
                  if (res.error) {
                    return reject(res.error);
                  }

                  Office.context.document.goToByIdAsync(currentSlideId, Office.GoToType.Slide, function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                      return reject("Action failed with error: " + asyncResult.error.message);
                    }

                    resolve();
                  });
                }
              );
            });
          });
        });
    });
  });
}
