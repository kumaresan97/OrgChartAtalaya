import * as React from "react";
import { useState, useEffect, useRef } from "react";
import Orgchart from "@balkangraph/orgchart.js";
import { NormalPeoplePicker } from "@fluentui/react/lib/Pickers";
import { MSGraphClient } from "@microsoft/sp-http";
import Styles from "./AtalayaOrgChart.module.scss";
import { IOrgChartProps, IpeoplePicker } from "./IOrgChart";
import { Spinner } from "@fluentui/react";
import "./style.css";

const azureAdUsers = [];

const OrgChart = (props: any): JSX.Element => {
  const [selectedpeoplePicker, setSelectedpeoplePicker] = useState<
    IpeoplePicker[]
  >([]);
  const [orgChartNode, setOrgChartNode] = useState([]);
  const [loader, setLoader] = useState<boolean>(false);
  const chartContainerRef = useRef(null);

  // get users from Asure
  async function getUsers() {
    setLoader(true);
    await props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("users")
          // .select(
          //   "department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones,officeLocation,identities,accountEnabled"
          // )
          .select(
            "onPremisesDistinguishedName,department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones,officeLocation,identities,accountEnabled"
          )
          .expand("manager")
          .top(999)
          .get()
          .then(async (data: any) => {
            azureAdUsers.push(...data.value);
            let strtoken = "";
            if (data["@odata.nextLink"]) {
              strtoken = data["@odata.nextLink"].split("skiptoken=")[1];
              await getnextUsers(
                data["@odata.nextLink"].split("skiptoken=")[1]
              );
            } else {
              await nodeBindHandler(azureAdUsers);
            }
          })
          .catch((error: any) => {
            console.log(error);
          });
      });
  }
  async function getnextUsers(skiptoken) {
    await props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("users")
          // .select(
          //   "department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones,officeLocation,identities,accountEnabled"
          // )
          .select(
            "onPremisesDistinguishedName,department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones,officeLocation,identities,accountEnabled"
          )
          // .select("*")
          .expand("manager")
          .skipToken(skiptoken)
          .get()
          .then(async (data: any) => {
            azureAdUsers.push(...data.value);
            let strtoken = "";
            if (data["@odata.nextLink"]) {
              strtoken = data["@odata.nextLink"].split("skipToken=")[1];
              await getnextUsers(
                data["@odata.nextLink"].split("skipToken=")[1]
              );
            } else {
              await nodeBindHandler(azureAdUsers);
            }
          })
          .catch((error: any) => {
            console.log(error);
          });
      });
  }

  let nodeBindHandler = (data: any) => {
    // data = Users.testdata();
    let orgChartUsers = [];
    data = data.filter(
      (e) =>
        e.onPremisesDistinguishedName !== null &&
        e.onPremisesDistinguishedName.includes("OU=Employees") &&
        e.onPremisesDistinguishedName.includes("OU=ATL") &&
        e.onPremisesDistinguishedName.includes("DC=atalayacap") &&
        e.onPremisesDistinguishedName.includes("DC=com")
    );
    data.forEach(
      (user: any) =>
        user.department !== null &&
        // (user.userPrincipalName.includes("atalayacap") ||
        //   user.userPrincipalName.includes("ATL")) &&
        orgChartUsers.push({
          id: user.id,
          pid: user?.manager?.id ? user?.manager?.id : null,
          title: user?.jobTitle ? user?.jobTitle : "N/A",
          manager: user?.manager?.displayName,
          department: user?.department,
          name: user?.displayName,
          email: user?.userPrincipalName ? user?.userPrincipalName : "N/A",
          img: `${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=${user?.userPrincipalName}`,
        })
    );
    setOrgChartNode([...orgChartUsers]);
    LoadFilteredChartData(orgChartUsers);
  };

  let LoadFilteredChartData = (
    input: IOrgChartProps[] = orgChartNode,
    searchName: string = "All"
  ) => {
    setLoader(true);
    const filtered: IOrgChartProps[] = [];
    const employee = input.find((employee) => employee.email === searchName);
    if (employee) {
      filtered.push({ ...employee, pid: "" });
      const stack = [employee.id];
      while (stack.length > 0) {
        const currentId = stack.pop();
        const children = input.filter((child) => child.pid === currentId);
        filtered.push(...children);
        stack.push(...children.map((child) => child.id));
      }
      loadChart(filtered);
    } else {
      loadChart(input);
    }
  };
  //  NormalPeoplePicker Function
  const GetUserDetails = (filterText: any) => {
    let peoples: IpeoplePicker[] = [];

    orgChartNode.map((user) =>
      // testuser.testdata().map((user) =>
      peoples.push({
        ID: user?.id,
        imageUrl: `${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=${user?.email}`,
        text: user?.name,
        secondaryText: user?.email,
      })
    );
    let result: any = peoples.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );

    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };
  const doesTextStartWith = (text: string, filterText: string): boolean => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };

  let loadChart = (_node: any) => {
    setLoader(false);
    var chart = new Orgchart(document.getElementById("tree"), {
      template: "olivia",
      mode: "light",
      layout: Orgchart.normal,
      // mouseScrool: Orgchart.action.ctrlZoom,
      mouseScrool: Orgchart.action.scroll,
      scaleInitial: 0.65,
      enableSearch: false,
      collapse: {
        level: 3,
      },
      toolbar: {
        layout: true,
        zoom: true,
        fit: true,
        expandAll: true,
      },
      menu: {
        pdf: { text: "Export PDF" },
        png: { text: "Export PNG" },
        svg: { text: "Export SVG" },
        csv: { text: "Export CSV" },
      },
      editForm: {
        generateElementsFromFields: false,
        elements: [
          { type: "textbox", label: "Name", binding: "name" },
          { type: "textbox", label: "Title", binding: "title" },
          { type: "textbox", label: "Department", binding: "department" },
          { type: "textbox", label: "Manager Name", binding: "manager" },
        ],
      },
      nodeBinding: {
        field_0: "name",
        field_1: "title",
        img_0: "img",
      },
      nodes: [..._node],
    });
    Orgchart.scroll.smooth = 2;
    Orgchart.scroll.speed = 25;
  };
  useEffect(() => {
    getUsers();
  }, []);
  return (
    <div className={Styles.orgchart_wraper}>
      <div
        // style={{ display: "flex", justifyContent: "end" }}
        className={Styles.filter_wraper}
      >
        <NormalPeoplePicker
          inputProps={{ placeholder: "Search User" }}
          onResolveSuggestions={GetUserDetails}
          itemLimit={1}
          selectedItems={selectedpeoplePicker}
          defaultSelectedItems={[
            { key: "All Deparments", text: "All Deparments" },
          ]}
          onChange={(selectedUser: any): void => {
            if (selectedUser.length) {
              setSelectedpeoplePicker([...selectedUser]);
              LoadFilteredChartData(
                orgChartNode,
                selectedUser[0].secondaryText
              );
            } else {
              setSelectedpeoplePicker([]);
              LoadFilteredChartData(orgChartNode);
            }
          }}
        />
      </div>
      {loader ? (
        <Spinner label="Loading..." />
      ) : (
        <div ref={chartContainerRef} id="tree" />
      )}
    </div>
  );
};

export default OrgChart;
