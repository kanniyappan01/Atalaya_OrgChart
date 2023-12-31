import * as React from 'react';
import Orgchart from '@balkangraph/orgchart.js';
import { MSGraphClient } from "@microsoft/sp-http";
import { useState, useEffect, useRef } from 'react';
import { NormalPeoplePicker } from '@fluentui/react';
import { IOrgChartProps, IpeoplePicker } from './IOrgChart';
import { Spinner } from '@fluentui/react';
import "./style.css";
// import testusers from './Users';
// import Styles from "./AtalayaOrgChart.module.scss";

// const dropdownStyles: Partial<IDropdownStyles> = {
//     dropdown: { width: 200 },
//     root: {
//         selectors: {
//             ".ms-Dropdown": {
//                 marginRight: 10,
//             },
//             ".ms-Dropdown-title": {
//                 height: 40,
//                 paddingTop: 4,
//             },
//             ".ms-Dropdown-caretDownWrapper": {
//                 top: 3,
//             },
//         },
//     },
// };


// let filterKeys = {
//     department: "",
//     selectedpeoplePicker: []
// }
// let CurrentUser :string = "";

var azureAdUsers: IOrgChartProps[] = [];

const OrgChart = (props: any): JSX.Element => {
    const chartContainerRef = useRef(null);
    const [users, setUsers] = useState<any[]>([])
    // const [loader, setLoader] = useState<boolean>(false)
    const [orgChartNode, setOrgChartNode] = useState<IOrgChartProps[]>([]);
    const [selectedpeoplePicker, setSelectedpeoplePicker] = useState<IpeoplePicker[]>([]);

    // depatment Filter

    // const [DepartmentDropdown, setDepartmentDropdown] = useState([]);
    // const removeDuplicates = (arr) => {
    //     return arr.filter((item, index, self) => {
    //         return index === self.findIndex((t) => (
    //             t.key === item.key // Replace 'id' with the property you want to check for uniqueness
    //         ));
    //     });
    // }
    // const createDrodownOptions = (data) => {
    //     let deparment = []
    //     // departmentDrapdown filter
    //     data.forEach((user, i) => user.department !== null && deparment.push({ key: user?.department, text: user?.department }));
    //     let dropdownValue = removeDuplicates(deparment)
    //     setDepartmentDropdown([{ key: "All Deparments", text: "All Deparments" }, ...dropdownValue])
    // }

    // const filterHandler = () => {

    //     let output = filterKeys.department !== "All Deparments" ? testusers.testdata().filter((user) => user?.department === filterKeys.department) :
    //         azureAdUsers
    //     console.log(output)
    //     setUsers(output)
    // }

    let nodeBindHandler = (data: any) => {
        // data = testusers.testdata();
        let orgChartUsers = [];
        data.forEach((user: any) =>
            // (user.accountEnabled == true || user.userType !== "Guest") &&
            orgChartUsers.push({
                id: user.id,
                pid: user?.manager?.id ? user?.manager?.id : null,
                title: user?.jobTitle ? user?.jobTitle : "N/A",
                manager: user?.manager?.displayName,
                department: user?.department,
                name: user?.displayName,
                email: user?.userPrincipalName ? user?.userPrincipalName : "N/A",
                img: `${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=${user?.userPrincipalName}`
            }))
        setOrgChartNode([...orgChartUsers]);
        LoadFilteredChartData(orgChartUsers);
        // createDrodownOptions(data)
    }

    let loadChart = (_node: any) => {
        // setLoader(false)
        console.log("loadChart",_node)
        Orgchart.templates.myTemplate = Object.assign({}, Orgchart.templates.olivia);
        Orgchart.templates.myTemplate.field_0 = '<text data-width="230" data-text-overflow="ellipsis"  style="font-size: 24px;" fill="#757575" x="125" y="100" text-anchor="middle">{val}</text>';
        Orgchart.templates.myTemplate.field_1 = '<text data-width="230" data-text-overflow="multiline" style="font-size: 16px;" fill="#757575" x="125" y="30" text-anchor="middle">{val}</text>';

        var chart = new Orgchart(document.getElementById("tree"), {
            mode: 'light',
            template: "olivia",
            // layout: Orgchart.tree,
            layout: Orgchart.treeRightOffset,
            scaleInitial: 0.75,
            enableSearch: false,
            mouseScrool: Orgchart.action.scroll,
            toolbar: {
                layout: true,
                zoom: true,
                fit: true,
                expandAll: true
            },
            editForm: {
                // readOnly: true,
                generateElementsFromFields: false,
                elements: [
                    { type: 'textbox', label: 'Name', binding: 'name' },
                    { type: 'textbox', label: 'Title', binding: 'title' },
                    { type: 'textbox', label: 'Department', binding: 'department' },
                    { type: 'textbox', label: 'Manager Name', binding: 'manager' }
                ]
            },
            nodeBinding: {
                field_0: "name",
                field_1: "title",
                img_0: "img"
            },
            nodes: [..._node],
        });
    }
    let LoadFilteredChartData = (input: IOrgChartProps[] = orgChartNode, searchName: string = "All") => {
        // setLoader(true)
        const filtered: IOrgChartProps[] = [];
        const employee = input.find(employee => employee.email === searchName);
        if (employee) {
            filtered.push({ ...employee, pid: '' });
            const stack = [employee.id];
            while (stack.length > 0) {
                const currentId = stack.pop();
                const children = input.filter(child => child.pid === currentId);
                filtered.push(...children);
                stack.push(...children.map(child => child.id));
            }
            // console.log("filtered", filtered)
            loadChart(filtered)
            console.log("filtered",filtered)
        } else {
            loadChart(input)
            console.log("non filtered",input)
        }
    };

    //  NormalPeoplePicker Function
    const GetUserDetails = (filterText: any) => {
        let peoples: IpeoplePicker[] = []

        users.map((user) => peoples.push({ ID: user?.id, imageUrl: `${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=${user?.userPrincipalName}`, text: user?.displayName, secondaryText: user?.userPrincipalName }))
        // testusers.testdata().map((user) => peoples.push({ ID: user?.id, imageUrl: `${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=${user?.userPrincipalName}`, text: user?.displayName, secondaryText: user?.userPrincipalName }))
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

    // get users from Asure
    async function  getUsers () {
        // setLoader(true)
        await props.context.msGraphClientFactory
            .getClient()
            .then((client: MSGraphClient) => {
                client
                    .api("users")
                    .select(
                        "department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones,officeLocation,identities,accountEnabled"
                    )
                    .expand("manager")
                    .top(999)
                    .get()
                    .then((data: any) => {
                        azureAdUsers.push(...data.value)
                        let strtoken = "";
                        if (data["@odata.nextLink"]) {
                            strtoken = data["@odata.nextLink"].split("skiptoken=")[1];
                            getnextUsers(data["@odata.nextLink"].split("skiptoken=")[1]);
                        } else {
                            setUsers(azureAdUsers)
                            nodeBindHandler(azureAdUsers);
                        }
                    })
                    .catch((error: any) => {
                        console.log(error);
                    });
            });
    }
    async function getnextUsers (skiptoken) {
        await props.context.msGraphClientFactory
            .getClient()
            .then((client: MSGraphClient) => {
                client
                    .api("users")
                    .select(
                        "department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones,officeLocation,identities,accountEnabled"
                    )
                    .expand("manager")
                    .top(999)
                    .skipToken(skiptoken)
                    .get()
                    .then((data: any) => {
                        azureAdUsers.push(...data.value)
                        let strtoken = "";
                        if (data["@odata.nextLink"]) {
                            strtoken = data["@odata.nextLink"].split("skipToken=")[1];
                            getnextUsers(data["@odata.nextLink"].split("skipToken=")[1]);
                        } else {
                            setUsers(azureAdUsers)
                            nodeBindHandler(azureAdUsers);
                        }
                    })
                    .catch((error: any) => {
                        console.log(error);
                    });
            });
    }

    // get current user
    // async function getCurrnetUser() {
    //     await props.context.msGraphClientFactory
    //         .getClient()
    //         .then((client: MSGraphClient) => {
    //             client
    //                 .api("me")
    //                 .select(
    //                     "department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones,officeLocation,identities"
    //                 )
    //                 .expand("manager")
    //                 .top(999)
    //                 .get()
    //                 .then(function (data) {
    //                     CurrentUser = data.userPrincipalName;

    //                     getUsers();
    //                 })
    //                 .catch(function (error) {
    //                     console.log(error);
    //                 });
    //         });

    // }

    useEffect(() => {
        getUsers();
    }, [])
    return (
        <div >
            <div style={{ display: "flex", justifyContent: "end" }}>
                {/* <Dropdown
                    placeholder="All Deparments"
                    selectedKey={filterKeys.department}
                    options={DepartmentDropdown}
                    styles={dropdownStyles}
                    onChange={(event, option: any) => {
                        setSelectedpeoplePicker([])
                        filterKeys.selectedpeoplePicker = []
                        filterKeys.department = option.key;
                        LoadFilteredChartData(orgChartNode)
                        filterHandler()
                    }}
                /> */}
                <button onClick={()=>loadChart(orgChartNode)}>All</button>
                <NormalPeoplePicker
                    inputProps={{ placeholder: "Search User" }}
                    onResolveSuggestions={GetUserDetails}
                    itemLimit={1}
                    selectedItems={selectedpeoplePicker}
                    defaultSelectedItems={[{ key: "All Deparments", text: "All Deparments" }]}
                    onChange={(selectedUser: any): void => {
                        if (selectedUser.length) {
                            setSelectedpeoplePicker([...selectedUser]);
                            // filterKeys.selectedpeoplePicker = [...selectedUser]
                            LoadFilteredChartData(orgChartNode, selectedUser[0].secondaryText);
                        } else {
                            setSelectedpeoplePicker([]);
                            // filterKeys.selectedpeoplePicker = []
                            LoadFilteredChartData(orgChartNode)
                        }
                    }}
                />
            </div>
            {/* {loader ? <Spinner label="Loading..." /> : } */}
            <div ref={chartContainerRef} id='tree' />
        </div>
    )
}

export default OrgChart