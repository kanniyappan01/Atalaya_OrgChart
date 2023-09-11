import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import Orgchart from '@balkangraph/orgchart.js';
import { graph } from '@pnp/graph/presets/all';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import Styles from "./AtalayaOrgChart.module.scss"
import "./style.css"

// var orgChartNode = []

const OrgChart = (props: any): JSX.Element => {
    const [users, setUsers] = useState([])
    const [userDatas, setUserDatas] = useState([]);
    const [orgChartNode, setOrgChartNode] = useState([])
    const chartContainerRef = useRef(null);
    let CurrentUser = props.context.pageContext.user.email;

    // geting users from Asure 
    let getUsers = async () => {
        await graph.users.top(999)
            .select(
                "department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones,officeLocation,identities"
            ).expand("manager")
            .get().then((response) => {
                setUsers(response)
                nodeBindHandler(response)
            }).catch((err) => {
                console.log(err);
            })
    }
    let nodeBindHandler = (data) => {
        
        
        let demoOutput= [];

        let demoNode = data.map((user,i)=>demoOutput.push({
            id:user.id,
            pid: user?.manager?.id ? user?.manager?.id:null,
            title: user?.jobTitle ? user?.jobTitle : "N/A",
            manager:user?.manager?.displayName,
            department: user?.department,
            name:user?.displayName,
            email: user?.userPrincipalName ? user?.userPrincipalName : "N/A",
            img: `${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=${user?.userPrincipalName}`
        }))
        console.log(demoOutput)
        setOrgChartNode([...demoOutput])
        LoadFilteredChartData(demoOutput)
    }

    let loadChart = (_node) => {
        Orgchart.templates.myTemplate = Object.assign({}, Orgchart.templates.olivia);
        Orgchart.templates.myTemplate.field_0 = '<text data-width="230" data-text-overflow="ellipsis"  style="font-size: 24px;" fill="#757575" x="125" y="100" text-anchor="middle">{val}</text>';
        Orgchart.templates.myTemplate.field_1 = '<text data-width="230" data-text-overflow="multiline" style="font-size: 16px;" fill="#757575" x="125" y="30" text-anchor="middle">{val}</text>';

        var chart = new Orgchart(document.getElementById("tree"), {
            mode: 'light',
            template: "olivia",
            layout: Orgchart.tree,
            scaleInitial: 0.75,
            enableSearch: false,
            mouseScrool: Orgchart.action.scroll,
            toolbar:{
                layout:true,
                zoom:true,
                fit:true,
                expandAll:true
            },
            editForm: {
                
                // readOnly: true,
                generateElementsFromFields: false,
                elements: [
                    { type: 'textbox', label: 'Name', binding: 'name'},
                    { type: 'textbox', label: 'Title', binding: 'title'},   
                    { type: 'textbox', label: 'Department', binding: 'department'},  
                    { type: 'textbox', label: 'Manager Name', binding: 'manager'}   
                ]
            },
            nodeBinding: {
                field_0: "name",
                field_1: "title",
                img_0: "img"
            },
            nodes: [..._node],
        });
        chart.on('click',(event, node)=>{
            console.log('click value:',event)
        });

    }
    let LoadFilteredChartData = (input=orgChartNode, searchName=CurrentUser) => {
        const filtered = [];
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
            console.log(filtered)
            loadChart(filtered)
        } else {
            loadChart(input)
        }
    };
    
    //  NormalPeoplePicker Function
    const GetUserDetails = (filterText: any) => {
        let peoples = []
        users.map((user) => peoples.push({ ID: user?.id, imageUrl: `${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=${user?.mail}`, text: user?.displayName, secondaryText: user?.mail }))
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
    
    useEffect(() => {
        getUsers();
    }, [])
    return (
        <div className={Styles.orgchart_wraper}>
            <NormalPeoplePicker
                inputProps={{ placeholder: "Search User" }}
                onResolveSuggestions={GetUserDetails}
                itemLimit={1}
                className={Styles.filter_wraper}
                selectedItems={userDatas}
                onChange={(selectedUser: any): void => {
                    if (selectedUser.length) {
                        setUserDatas([...selectedUser]);
                        LoadFilteredChartData(orgChartNode, selectedUser[0].secondaryText);
                    } else {
                        setUserDatas([]);
                        LoadFilteredChartData(orgChartNode)
                    }
                }}
            />
            <div ref={chartContainerRef} id='tree' />
        </div>
    )
}

export default OrgChart