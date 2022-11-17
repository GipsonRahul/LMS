import * as React from 'react';
import { IUngotiApplyLeaveProps } from './IUngotiApplyLeaveProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { createStyles, makeStyles, Theme } from '@material-ui/core/styles';
import styles from "./UngotiApplyLeave.module.scss";
import { useRef, useState } from 'react';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";
import { forwardRef } from 'react';
import { IUngotiApplyLeaveState, LeaveDetails } from './IUngotiApplyLeaveState';

import Grid from '@material-ui/core/Grid';
import DeleteIcon from '@material-ui/icons/Delete';
import VisibilityIcon from '@material-ui/icons/Visibility';
import AddIcon from '@material-ui/icons/Add';
import InputLabel from '@material-ui/core/InputLabel';
import FormHelperText from '@material-ui/core/FormHelperText';
import FormControl from '@material-ui/core/FormControl';
import Select from '@material-ui/core/Select';

import Button from '@material-ui/core/Button';
import Modal from '@material-ui/core/Modal';
import TextField from '@material-ui/core/TextField';
import Dialog from '@material-ui/core/Dialog';
import DialogActions from '@material-ui/core/DialogActions';
import DialogContent from '@material-ui/core/DialogContent';
import DialogContentText from '@material-ui/core/DialogContentText';
import DialogTitle from '@material-ui/core/DialogTitle';

import IconButton from '@material-ui/core/IconButton';
import PhotoCamera from '@material-ui/icons/PhotoCamera';
import AttachmentIcon from '@material-ui/icons/Attachment';
import AddCircleOutlineIcon from '@material-ui/icons/AddCircleOutline';
import {
  Typography, ButtonGroup,
  ListItem,
  Badge,
  List,
  ListItemText,
  LinearProgress,
  Menu
} from '@material-ui/core';

import ArrowDropDownIcon from '@material-ui/icons/ArrowDropDown';
import MenuItem from '@material-ui/core/MenuItem';


import MaterialTable, { Column, Icons } from 'material-table';

import AssignmentTurnedInIcon from '@material-ui/icons/AssignmentTurnedIn';
import EditIcon from '@material-ui/icons/Edit';

import AddBox from '@material-ui/icons/AddBox';
import ArrowDownward from '@material-ui/icons/ArrowDownward';
import Check from '@material-ui/icons/Check';
import ChevronLeft from '@material-ui/icons/ChevronLeft';
import ChevronRight from '@material-ui/icons/ChevronRight';
import Clear from '@material-ui/icons/Clear';
import DeleteOutline from '@material-ui/icons/DeleteOutline';
import Edit from '@material-ui/icons/Edit';
import FilterList from '@material-ui/icons/FilterList';
import FirstPage from '@material-ui/icons/FirstPage';
import LastPage from '@material-ui/icons/LastPage';
import Remove from '@material-ui/icons/Remove';
import SaveAlt from '@material-ui/icons/SaveAlt';
import Search from '@material-ui/icons/Search';
import ViewColumn from '@material-ui/icons/ViewColumn';
import '../../../scss/styles.scss';
import "alertifyjs";
var alertify: any = require("../../../ExternalRef/JS/alertify.min.js");

export interface IHRProps {
  graphClient: any;
  currentContext:any;
}
var listUrl="";
export default class HR extends React.Component<IHRProps, any> {
  public deleteId = 0;

  public oldLeaveTypeId = 0;
  public oldNoofDays = 0;
  public txtSelectDate = 'Select Date';

  public file = null;

  constructor(props) {
    super(props);

    this.state = {
      formData: {},
      listLeaveDetails: [],
      copyListLeaveDetails: [],
      listLeaveBalance: [],
      allLeaveTypes: [],
      departments: [],
      allUserwithDepartments: [],
      fileDetails:[],
      disableresponse:false,
      disableAdddays:false,
      Requester:""

    };

    alertify.set("notifier", "position", "top-center");
    listUrl = this.props.currentContext.pageContext.web.absoluteUrl;
    var siteindex = listUrl.toLocaleLowerCase().indexOf("sites");
    var teamsindex= listUrl.toLocaleLowerCase().indexOf("teams");
    if(siteindex>=0)
    {
      listUrl = listUrl.substr(siteindex - 1) + "/Lists/";
    }
    else  if(teamsindex>=0)
    {
      listUrl = listUrl.substr(siteindex - 1) + "/Lists/";
    }
    else
    {
      listUrl = "/Lists/";
    }

    this.loadDepartment();

    sp.web.currentUser.get().then((userdata) => {
      this.setState({ currentUser: userdata }); 
      this.loadLeaveTypes(); 
    });  

  }
  
  public loadDepartment = () => {
    this.props.graphClient
      .api("/users?$select=Department,mail")  
      .get()
      .then((response: any) => {
        var departments = this.state.departments;
        departments = ['Select'];
        var allUserwithDepartments = response.value;
        for (let index = 0; index < allUserwithDepartments.length; index++) {
          const value = allUserwithDepartments[index];
          var hasdata = departments.filter(c => c == value.department);
          if (hasdata.length == 0) {
            departments.push(value.department);
          }
        }
        this.setState({ allUserwithDepartments: allUserwithDepartments, departments: departments });
      }).catch(error => {
        console.log(error);
      });
  }

  public loadLeaveTypes = () => {
    // sp.web.lists.getByTitle("LeaveTypes")
    sp.web.getList(listUrl + "LeaveTypes")
    
      .items
      .filter("Active eq '1'")
      .get()
      .then((res) => {
        var allLeaveTypes = this.state.allLeaveTypes;
        allLeaveTypes = [];
        var columns = this.state.columns;
        columns = [
          { field: 'EmployeeName', title: 'Employee' }
        ];
        for (let index = 0; index < res.length; index++) {
          const leaveType = res[index];
          allLeaveTypes.push({
            Id: leaveType.Id,
            Title: leaveType.Title,
            DisplayName: leaveType.ScreenName
          });
          columns.push(
            {
              field: leaveType.Title, title: leaveType.ScreenName
            },
            {
              field: leaveType.Title + 'Used', title: leaveType.ScreenName + ' Consumed'
            });
        }
        this.setState({ allLeaveTypes: allLeaveTypes, columns: columns });
        this.loadLeaveBalance();
      });
  }

  public loadLeaveBalance = () => {
    var select = 'Id,EmployeeEmail/Id,EmployeeEmail/Title,EmployeeEmail/EMail';
    for (let index = 0; index < this.state.allLeaveTypes.length; index++) {
      const leaveType = this.state.allLeaveTypes[index].Title;
      select = select + ',' + leaveType + ',' + leaveType + 'Used';
    }
    var currentYear = new Date().getFullYear();
    // sp.web.lists.getByTitle("LeaveBalance")
    sp.web.getList(listUrl + "LeaveBalance")

      .items
      // .filter("Year eq '" + currentYear + "' and EmployeeEmailId ne '" + this.state.currentUser.Id + "'")
      .filter("Year eq '" + currentYear + "'")
      .select(select)
      .expand("EmployeeEmail")
      .get()
      .then((res) => {
        var leaveBalance = this.state.listLeaveBalance;
        leaveBalance = [];
        if (res.length > 0) {
          for (let index = 0; index < res.length; index++) {
            var data = res[index];
            var result = {
              Id: data.Id,
              EmployeeName: data.EmployeeEmail.Title,
              EMail: data.EmployeeEmail.EMail
            };
            for (let j = 0; j < this.state.allLeaveTypes.length; j++) {
              const leaveType = this.state.allLeaveTypes[j];
              result[leaveType.Title] = data[leaveType.Title];
              result[leaveType.Title + 'Used'] = data[leaveType.Title + 'Used'];
            }
            leaveBalance.push(result);
          }
        }
        this.setState({ listLeaveBalance: leaveBalance, copyListLeaveBalance: leaveBalance });
      });

      this.loadAppliedLeave();
  }

  public loadAppliedLeave = () => {
    // sp.web.lists.getByTitle("LeaveRequest")
    sp.web.getList(listUrl + "LeaveRequest")

      .items
      .select("Id", "Title", "From", "To", "NoofDays","Created", "Detail", "Status", "LeaveType/Id", "LeaveType/Title", "LeaveType/ScreenName", "ApproverId","Requester/Title", "Requester/Id", "Requester/FirstName", "Requester/EMail")
      .filter("Status eq 'Pending'")
      .orderBy('Modified', false)
      .expand("LeaveType", "Requester")
      .get()
      .then((res) => {
        var lstData = this.state.listLeaveDetails;
        lstData = [];
        for (let index = 0; index < res.length; index++) {
          const leave = res[index];
          lstData.push({
            Id: leave.Id,
            ApproverId: leave.ApproverId,
            RequesterId: leave.Requester.Id,
            RequesterFirstName: leave.Requester.Title,
            LeaveTypeTitle: leave.LeaveType.Title,
            LeaveType: leave.LeaveType.ScreenName,
            LeaveTypeId: leave.LeaveType.Id,
            From: new Date(leave.From),
            strFrom: this.formatDate(new Date(leave.From)),
            To: new Date(leave.To),
            strTo: this.formatDate(new Date(leave.To)),
            NoofDays: parseFloat(leave.NoofDays),
            strNoofDays: leave.NoofDays,
            Detail: leave.Detail,
            Status: leave.Status,
            CreatedDate:new Date(leave.Created),
            RequestedDate:this.formatDate(new Date(leave.Created)),
            RequesterMail:leave.Requester.EMail
          });
        }
        this.setState({ listLeaveDetails: lstData, copyListLeaveDetails: lstData });
      });
  }
  public formatDate = (paramdate) => {
    var date = paramdate.getDate() + '';
    if (date.length == 1) {
      date = '0' + date;
    }
    var month = (paramdate.getMonth() + 1) + '';
    if (month.length == 1) {
      month = '0' + month;
    }
    return date + "/" + month + "/" + paramdate.getFullYear();
  }


  public inputChangeHandler = (e) => {
    let Days = e.target.value;

    if (!Number(Days)) {
      this.setState({
        totalDays: ''
      });
        return;
    }
    this.setState({
      totalDays: e.target.value
    });
  }

  public setLeaveType = (event: React.ChangeEvent<any>) => {
    // this.setState({ selLeaveType: event.target.value, totalDays: this.state.formdata[event.target.value] });
    this.setState({ selLeaveType: event.target.value });
    if (event.target.value) {
      this.setState({ errorleavetype: null });
    }
  }

  public closePopup = () => {
    this.setState({ showUpdate: false ,disableAdddays:false});
  }

  public editLeaveBalance = (id,EmployeeName) => {

    // sp.web.lists.getByTitle("LeaveBalance")
    sp.web.getList(listUrl + "LeaveBalance")

      .items.getById(id)
      .get()
      .then((response) => {
        this.setState({ formData: response });
      });

    this.setState({ showUpdate: true, id: id,disableAdddays:false,EmployeeName: EmployeeName});
  }

  public filterbyDepartment = (event: React.ChangeEvent<any>) => {
    if (event.target.value != 'Select') {
      var allLeaveBalance = this.state.copyListLeaveBalance;
      var allUserwithDepartments = this.state.allUserwithDepartments;
      var departmentUsers = allUserwithDepartments.filter(c => c.department == event.target.value);
      var filteredLeaveBalance = this.state.listLeaveBalance;
      filteredLeaveBalance = [];
      for (let i = 0; i < departmentUsers.length; i++) {
        const mail = departmentUsers[i].mail;
        var leaveBalances = allLeaveBalance.filter(c => c.EMail == mail);
        filteredLeaveBalance = filteredLeaveBalance.concat(leaveBalances);
      }
      this.setState({ listLeaveBalance: filteredLeaveBalance });
    } else {
      this.setState({ listLeaveBalance: this.state.copyListLeaveBalance });
    }
  }
  public filterLeavebyDepartment = (event: React.ChangeEvent<any>) => {
    if (event.target.value != 'Select') {
      var allLeaveDetails = this.state.copyListLeaveDetails;
    
      var allUserwithDepartments = this.state.allUserwithDepartments;
      var departmentUsers = allUserwithDepartments.filter(c => c.department == event.target.value);
      // var filteredLeaveBalance = this.state.listLeaveBalance;
      var filteredLeaveDetails = [];
      for (let i = 0; i < departmentUsers.length; i++) {
        const mail = departmentUsers[i].mail;
        var finalleaveDetails = allLeaveDetails.filter(c => c.RequesterMail == mail);
        filteredLeaveDetails = filteredLeaveDetails.concat(finalleaveDetails);
      }
      this.setState({ listLeaveDetails: filteredLeaveDetails });
    } else {
      this.setState({ listLeaveDetails: this.state.copyListLeaveDetails });
    }
  }
  

  public submit = () => {
    if (!this.state.selLeaveType) {
      this.setState({ errorleavetype: 'Leave type is required' });
      return;
    }
    if (!this.state.totalDays) {
      this.setState({ errorNoofdays: 'No of days is required' });
      return;
    }
    this.setState({disableAdddays:true})
    var formData = this.state.formData;
    var currentDays=formData[this.state.selLeaveType];
    var addedtotalDays=Number(currentDays)+Number(this.state.totalDays);
    formData[this.state.selLeaveType] = addedtotalDays.toString();
    // sp.web.lists.getByTitle("LeaveBalance")
    sp.web.getList(listUrl + "LeaveBalance")

      .items.getById(this.state.id)
      .update(formData)
      .then((response) => {
        alertify.success('Leave balance updated');
        this.setState({
          totalDays: "",
          errorleavetype:'',
          errorNoofdays:'',
          selLeaveType:''
        });
        this.loadLeaveBalance();
        this.closePopup();
      });
  }

  public openConfirm = (value, id = 0, status = '',requesterName) => {
    this.setState({Requester:requesterName});
    if (value) {
      var currentYear = new Date().getFullYear();
      var currentleaveDetails = this.state.listLeaveDetails.filter(c => c.Id == id)[0];
      // sp.web.lists.getByTitle("LeaveBalance")
      sp.web.getList(listUrl + "LeaveBalance")

        .items
        .filter("Year eq '" + currentYear + "' and EmployeeEmailId eq '" + currentleaveDetails.RequesterId + "'")
        .get()
        .then((res) => {
          this.setState({ leaveBalance: res[0] });
        });
      this.getLeaveData(id, false);
      this.setState({ currentleaveDetails: currentleaveDetails,disableresponse:false });
    }
    this.setState({ showConfirm: value, id: id, status: status,disableresponse:false });
  }

  public getLeaveData = (id, view) => {
    // sp.web.lists.getByTitle("LeaveRequest")
    sp.web.getList(listUrl + "LeaveRequest")

      .items.getById(id).select("Id","From","To","NoofDays","Detail","Status","FromHalf","ToHalf","DocumentUrl","Approver/Id","Approver/Title","LeaveType/Id","Requester/Id","Requester/Title")
      .expand("LeaveType","Requester","Approver").
      get()
      .then((response) => {
        var from = new Date(response.From);
        var to = new Date(response.To);
        var formData = this.state.formData;
        formData = {
          Id: response.Id,
          ApproverId: response.ApproverId,
          RequesterId: response.RequesterId,
          LeaveTypeId: response.LeaveTypeId,
          From: from,
          To: to,
          NoofDays: response.NoofDays,
          Detail: response.Detail,
          Status: response.Status,
          FromHalf: response.FromHalf,
          ToHalf: response.ToHalf,
          DocumentUrl: response.DocumentUrl
        };
        var fromDefaultDate = this.formatDate(from);
        var toDefaultDate = this.formatDate(to);
        this.oldLeaveTypeId = response.LeaveTypeId;
        this.oldNoofDays = parseFloat(response.NoofDays);

        var name = '';

        if (response.DocumentUrl) {
          var docResponse=response.DocumentUrl.split(';');
          var docResponseFiles=[];
          for(let i=0;i<docResponse.length;i++)
          {
            if(docResponse[i])
            {
              docResponseFiles=this.state.fileDetails;
              this.file = docResponse[i];
              var sdata =  docResponse[i].split('/');
              name = sdata[sdata.length - 1];
              docResponseFiles.push({filenname:name,files:this.file});
            }
          }

        }

        if(view)
        {
          this.setState({ formData: formData, strFrom: fromDefaultDate, strTo: toDefaultDate, fileDetails: docResponseFiles, openAddPopup: false,openEditPopup:false, isview: view });

        }
        else{
          this.setState({ formData: formData, strFrom: fromDefaultDate, strTo: toDefaultDate, fileDetails: docResponseFiles, openAddPopup: false,openEditPopup:false, isview: false });
        }
      });
  }
  public viewLeave = (id,requesterfullname) => {
    this.setState({Requester:requesterfullname});
    this.getLeaveData(id, true);
  }

  public closeViewPopup = () => {
    this.setState({ isview: false });
  }

  public ApproveRequest = () => {
    if (this.state.status == 'Rejected' && !this.state.note) {
      this.setState({ errornote: 'Detail is required', _errornote: true });
      return;
    } else {
      this.setState({ errornote: '', _errornote: false });
    }
    this.setState({disableresponse:true});
    var formData = this.state.formData;
    var leaveBalance = this.state.leaveBalance;
    formData.ManagerNotes = this.state.note;
    formData.Status = this.state.status;
    var leaveTypeUsed = this.state.currentleaveDetails.LeaveTypeTitle + 'Used';
    var leaveTypePendingApproval = this.state.currentleaveDetails.LeaveTypeTitle + 'PendingApproval';
    leaveBalance[leaveTypePendingApproval] = (parseInt(leaveBalance[leaveTypePendingApproval]) - parseInt(formData.NoofDays)) + '';
    if (this.state.status == 'Approved') {
      leaveBalance[leaveTypeUsed] = (parseInt(leaveBalance[leaveTypeUsed]) + parseInt(formData.NoofDays)) + '';
    }
    // sp.web.lists.getByTitle("LeaveRequest")
    sp.web.getList(listUrl + "LeaveRequest")

      .items.getById(formData.Id)
      .update(formData)
      .then((response) => {
        this.updateLeaveBalance(leaveBalance);
        alertify.success('Leave updated successfully');
        this.loadAppliedLeave();
        this.setState({ showConfirm: false ,note: ''});
      });

  }

  public updateLeaveBalance = (leaveBalance) => {
    // sp.web.lists.getByTitle("LeaveBalance")
    sp.web.getList(listUrl + "LeaveBalance")
      .items.getById(leaveBalance.Id)
      .update(leaveBalance)
      .then((response) => {
      });
  }

  
  public inputChangeNotesHandler = (e) => {
    this.setState({
      note: e.target.value
    });
    if (e.target.value) {
      this.setState({ errornote: '', _errornote: false });
    }
  }

  public render(): React.ReactElement {

    const columns = [
      { field: 'LeaveType', title: 'Type' },
      { field: 'RequestedDate', title: 'Requested Date' },
      { field: 'RequesterFirstName', title: 'Requester' },
      { field: 'strFrom', title: 'From' },
      { field: 'strTo', title: 'To' },
      { field: 'strNoofDays', title: 'No. of days' },
      // { field: 'Status', title: 'Status' },
      // { field: 'Action', title: 'Action' },
    ];

    const tableIcons: Icons = {
      Add: forwardRef((props: any, ref: any) => <AddBox {...props} ref={ref} />),
      Check: forwardRef((props: any, ref: any) => <Check {...props} ref={ref} />),
      Clear: forwardRef((props: any, ref: any) => <Clear {...props} ref={ref} />),
      Delete: forwardRef((props: any, ref: any) => <DeleteOutline {...props} ref={ref} />),
      DetailPanel: forwardRef((props: any, ref: any) => <ChevronRight {...props} ref={ref} />),
      Edit: forwardRef((props: any, ref: any) => <Edit {...props} ref={ref} />),
      Export: forwardRef((props: any, ref: any) => <SaveAlt {...props} ref={ref} />),
      Filter: forwardRef((props: any, ref: any) => <FilterList {...props} ref={ref} />),
      FirstPage: forwardRef((props: any, ref: any) => <FirstPage {...props} ref={ref} />),
      LastPage: forwardRef((props: any, ref: any) => <LastPage {...props} ref={ref} />),
      NextPage: forwardRef((props: any, ref: any) => <ChevronRight {...props} ref={ref} />),
      PreviousPage: forwardRef((props: any, ref: any) => <ChevronLeft {...props} ref={ref} />),
      ResetSearch: forwardRef((props: any, ref: any) => <Clear {...props} ref={ref} />),
      Search: forwardRef((props: any, ref: any) => <Search {...props} ref={ref} />),
      SortArrow: forwardRef((props: any, ref: any) => <ArrowDownward {...props} ref={ref} />),
      ThirdStateCheck: forwardRef((props: any, ref: any) => <Remove {...props} ref={ref} />),
      ViewColumn: forwardRef((props: any, ref: any) => <ViewColumn {...props} ref={ref} />)
    };

    return (
      <div>
        <Grid container>
          <Grid sm={3} >
           
            <FormControl variant="outlined" className="form-group bg-white" size="small">
              <InputLabel id="standard-select-currency" >Department</InputLabel>
              <Select
                labelId="standard-select-currency"
                id="standard-select-currency"
                defaultValue="Select"
                onChange={this.filterbyDepartment}
                label="Departments"
              >
                {
                  this.state.departments.map((department) => {
                    return (
                      <MenuItem value={department}>{department}</MenuItem>
                    );
                  })
                }
              </Select>
            </FormControl>
          </Grid>
        </Grid>

        <div className="mt-25 manageLeave">
        <MaterialTable
          title="Leave Balance"
          icons={tableIcons}
          columns={this.state.columns}
          data={this.state.listLeaveBalance}
          actions={[
            (rowData: any) => ({
              icon: forwardRef((props: any, ref: any) => <AddCircleOutlineIcon />),
              tooltip: 'Add',
              onClick: (event, value) => this.editLeaveBalance(rowData.Id,rowData.EmployeeName),
            })
          ]}
          options={{
            exportButton: true
          }}
        />
        </div>


        <Dialog open={this.state.showUpdate} className="applyLeaveDialog">
          <DialogTitle className="MuiDialogTitle-bg" id="form-dialog-title">
        <Typography component={"h5"}>Add Leave Balance for {this.state.EmployeeName}</Typography>
          </DialogTitle>
          <DialogContent>
            <Grid container>
              <Grid sm={12} >
                <FormControl variant="outlined" className="form-group" size="small" error={this.state.errorleavetype ? true : false}>
                  <InputLabel id="standard-select-currency" >Leave Type</InputLabel>
                  <Select
                    labelId="standard-select-currency"
                    id="standard-select-currency"
                    value={this.state.LeaveTypeId} onChange={this.setLeaveType}
                    label="Leave Type"
                  >
                    {
                      this.state.allLeaveTypes.map((leaveType) => {
                        return (
                          <MenuItem value={leaveType.Title}>{leaveType.DisplayName}</MenuItem>
                        );
                      })
                    }
                  </Select>
                </FormControl>
                <FormHelperText>{this.state.errorleavetype}</FormHelperText>

              </Grid>

              <Grid sm={12} >
                <TextField
                  id="standard-select-currency"
                  error={this.state.errorNoofdays ? true : false}
                  value={this.state.totalDays}
                  onChange={(e) => this.inputChangeHandler.call(this, e)}
                  multiline
                  label="Days"
                  name="Days"
                  placeholder="Please enter no. of days"
                  className="form-group"
                  variant={"outlined"}
                  size={"small"}
                  type="number" 
                >
                </TextField>
              </Grid>
            </Grid>


          </DialogContent>
          <DialogActions>
            <Button variant="contained" disableElevation color="default" size="small" onClick={this.closePopup}>
              Cancel
          </Button>
            <Button variant="contained" disabled={this.state.disableAdddays} disableElevation color="primary" size="small" onClick={this.submit}>
              Apply
          </Button>

          </DialogActions>
        </Dialog>

        <div className={styles.ungotiApplyLeave}>
        <div >
        <Grid container>
          <Grid sm={3} >
           
            <FormControl variant="outlined" className="form-group bg-white" size="small">
              <InputLabel id="standard-select-currency" >Department</InputLabel>
              <Select
                labelId="standard-select-currency"
                id="standard-select-currency"
                defaultValue="Select"
                onChange={this.filterLeavebyDepartment}
                label="Departments"
              >
                {
                  this.state.departments.map((department) => {
                    return (
                      <MenuItem value={department}>{department}</MenuItem>
                    );
                  })
                }
              </Select>
            </FormControl>
          </Grid>
        </Grid>
          <section className="mt-25">
            <Grid container spacing={2}>

              {
                <Grid className="manageLeave" item xs={12} sm={12} md={12} lg={12} xl={12}>

                  <MaterialTable
                    title="Pending Leave"
                    icons={tableIcons}
                    columns={columns}
                    data={this.state.listLeaveDetails}
                    actions={[
                      (rowData: LeaveDetails) => ({
                        icon: forwardRef((props: any, ref: any) => <AssignmentTurnedInIcon />),
                        tooltip: 'Approve',
                        onClick: (event, value) => this.openConfirm(true, rowData.Id, 'Approved',rowData.RequesterFirstName),
                      }),
                      (rowData: LeaveDetails) => ({
                        icon: forwardRef((props: any, ref: any) => <DeleteIcon />),
                        tooltip: 'Reject',
                        onClick: (event, value) => this.openConfirm(true, rowData.Id, 'Rejected',rowData.RequesterFirstName),
                      }),  
                      {
                        icon: forwardRef((props: any, ref: any) => <VisibilityIcon />),
                        tooltip: 'View',
                        onClick: (event, rowData: LeaveDetails) => this.viewLeave(rowData.Id,rowData.RequesterFirstName),
                      }
                    ]} 
                    options={{
                      actionsColumnIndex: 6
                    }}
                  />

                </Grid>
              }


            </Grid>
          </section>
        </div>


        <Dialog open={this.state.isview} className="applyLeaveDialog">
          <DialogTitle className="MuiDialogTitle-bg" id="form-dialog-title">
            <Typography component={"h5"}>Leave Details for {this.state.Requester}</Typography>
          </DialogTitle>
          <DialogContent>
            <section className="dateRangePicker">
              <Grid container spacing={2} className="datefield">
                <Grid sm={12} md={6}>
                  <Typography component={"p"} className="small-text">
                    FROM DATE
               </Typography>
                  <Typography component={"p"}>
                    {this.state.strFrom}
                  </Typography>
                </Grid>
                <div className="dateRangePicker-totalDays">
                  <span className="number">{this.state.formData.NoofDays} Day(s)</span>
                </div>
                <Grid sm={12} md={6} className={"text-right"}>
                  <Typography component={"p"} className="small-text">
                    TO DATE
               </Typography>
                  <Typography component={"p"}>
                    {this.state.strTo}
                  </Typography>
                </Grid>
              </Grid>
            </section>
            <Grid container>
              <Grid sm={12} >
                <h2>Note</h2>
                <DialogContentText>
                  {this.state.formData.Detail}
                </DialogContentText>
              </Grid>

              <Grid sm={12} >

              </Grid>
            </Grid>

            {
                <Grid container spacing={2}>
                  <Grid item xs={12} sm={12} md={6} lg={6} xl={6} className="form-group">
                    {/* <label htmlFor="icon-button-file" className="uploadbtn">
                      <IconButton color="primary" aria-label="upload picture" component="span">
                        <AttachmentIcon />
                      </IconButton>
                      <label><a href={this.state.formData.DocumentUrl} target="_blank">{this.state.fileName}</a></label>
                    </label> */}
                     <h2>Documents</h2>
                    { 
              this.state.fileDetails ?
              this.state.fileDetails.map((filedet,index) => {
                           
                           return (
                             <>
                             <label><a href={filedet.files} target="_blank">{filedet.filenname}</a></label><br></br>
                           </>
                           );
                         
                         }):''

                        }

                      {/* <label htmlFor="icon-button-file" className="uploadbtn">
                        <Button color="primary" aria-label="upload picture" component="span" variant="contained" >
                          <AttachmentIcon />Documents
                        </Button>
                        </label> */}
                  </Grid>
                </Grid>
                
            }



          </DialogContent>
          <DialogActions>
            <Button variant="contained" disableElevation color="default" size="small" onClick={this.closeViewPopup}>
              Cancel
          </Button>
          </DialogActions>
        </Dialog>





        {/* <Dialog open={this.state.isview} className="applyLeaveDialog">
          <DialogTitle className="MuiDialogTitle-bg" id="form-dialog-title">
            <Typography component={"h5"}>Leave Details</Typography>
          </DialogTitle>
          <DialogContent>
            <section className="dateRangePicker">
              <Grid container spacing={2} className="datefield">
                <Grid sm={12} md={6}>
                  <Typography component={"p"} className="small-text">
                    FROM DATE
               </Typography>
                  <Typography component={"p"}>
                    {this.state.strFrom}
                  </Typography>
                </Grid>
                <div className="dateRangePicker-totalDays">
                  <span className="number">{this.state.formData.NoofDays} Day(s)</span>
                </div>
                <Grid sm={12} md={6} className={"text-right"}>
                  <Typography component={"p"} className="small-text">
                    TO DATE
               </Typography>
                  <Typography component={"p"}>
                    {this.state.strTo}
                  </Typography>
                </Grid>
              </Grid>
            </section>
            <Grid container>
              <Grid sm={12} >
                <h2>Reason</h2>
                <DialogContentText>
                  {this.state.formData.Detail}
                </DialogContentText>
              </Grid>

              <Grid sm={12} >

              </Grid>
            </Grid>

            {
              this.state.fileName ?
                <Grid container spacing={2}>
                  <Grid item xs={12} sm={12} md={6} lg={6} xl={6} className="form-group">
                    <label htmlFor="icon-button-file" className="uploadbtn">
                      <IconButton color="primary" aria-label="upload picture" component="span">
                        <AttachmentIcon />
                      </IconButton>
                      <label><a href={this.state.formData.DocumentUrl} target="_blank">{this.state.fileName}</a></label>
                    </label>
                  </Grid>
                </Grid>
                : ''
            }



          </DialogContent>
          <DialogActions>
            <Button variant="contained" disableElevation color="default" size="small" onClick={this.closeViewPopup}>
              Cancel
          </Button>
          </DialogActions>
        </Dialog> */}





        <Dialog open={this.state.showConfirm} className="applyLeaveDialog">
          <DialogTitle className="MuiDialogTitle-bg" id="form-dialog-title">
      <Typography component={"h5"}>Leave Request for {this.state.Requester}</Typography>
          </DialogTitle>
          <DialogContent>
            <Grid container>

              <Grid sm={12} >
                <TextField
                  id="standard-select-currency"
                  error={this.state._errornote}
                  helperText={this.state.errornote}
                  value={this.state.note}
                  onChange={(e) => this.inputChangeNotesHandler.call(this, e)}
                  multiline
                  label="Note"
                  name="Detail"
                  rows="4"
                  placeholder="Please enter any additional information"
                  className="form-group"
                  variant={"outlined"}
                  size={"small"}
                >
                </TextField>

              </Grid>
            </Grid>

          </DialogContent>
          <DialogActions>
            <Button variant="contained" disableElevation color="default" size="small" onClick={this.openConfirm.bind(this, false)}>
              Cancel
          </Button>
          {
            this.state.status == 'Rejected' ? <Button disabled={this.state.disableresponse} variant="contained" disableElevation color="primary" size="small" onClick={this.ApproveRequest}> Reject</Button>:<Button variant="contained" disabled={this.state.disableresponse} disableElevation color="primary" size="small" onClick={this.ApproveRequest}>
              Approve
          </Button>
          }


          </DialogActions>
        </Dialog>

      </div >


      </div>



    );
  }
}
