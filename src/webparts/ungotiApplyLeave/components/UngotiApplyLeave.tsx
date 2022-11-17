import * as React from 'react';
import { IUngotiApplyLeaveProps } from './IUngotiApplyLeaveProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { createStyles, makeStyles, Theme } from '@material-ui/core/styles';

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
import "@pnp/sp/profiles";

import Paper from '@material-ui/core/Paper';
import Table from '@material-ui/core/Table';
import TableBody from '@material-ui/core/TableBody';
import TableCell from '@material-ui/core/TableCell';
import TableContainer from '@material-ui/core/TableContainer';
import TableHead from '@material-ui/core/TableHead';
import TablePagination from '@material-ui/core/TablePagination';
import TableRow from '@material-ui/core/TableRow';

import Popper from '@material-ui/core/Popper';

import Grid from '@material-ui/core/Grid';
import DeleteIcon from '@material-ui/icons/Delete';
import EditIcon from '@material-ui/icons/Edit';
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

import { DateRangePicker } from "materialui-daterange-picker";

import { forwardRef } from 'react';

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

import MaterialTable, { Column, Icons } from 'material-table';


import IconButton from '@material-ui/core/IconButton';
import PhotoCamera from '@material-ui/icons/PhotoCamera';
import AttachmentIcon from '@material-ui/icons/Attachment';
import DeleteForeverIcon from '@material-ui/icons/DeleteForever';

import '../../../scss/styles.scss';

import Manager from "./Manager";
import HR from "./HR";
import Holidays from "./Holidays";

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
import ClickAwayListener from '@material-ui/core/ClickAwayListener';
import Grow from '@material-ui/core/Grow';
import MenuList from '@material-ui/core/MenuList';

import styles from "./UngotiApplyLeave.module.scss";

import "alertifyjs"; 

import "../../../ExternalRef/CSS/style.css"; 
import "../../../ExternalRef/CSS/alertify.min.css";
var alertify: any = require("../../../ExternalRef/JS/alertify.min.js");

import { IUngotiApplyLeaveState, LeaveDetails } from './IUngotiApplyLeaveState';

var folderPath = 'Leave Documents';
var currentDate = new Date(new Date().toDateString());
var azureGroupId = '3acdf24e-e034-4b68-bfbe-e717aed7219c';
var listUrl="";

export default class UngotiApplyLeave extends React.Component<IUngotiApplyLeaveProps, IUngotiApplyLeaveState> {

  public deleteId = 0;
  public managerId = 0;

  public oldLeaveTypeId = 0;
  public oldNoofDays = 0;
  public txtSelectDate = 'Select Date';

  // public file = null;
  public file = [];
  public leaveColors = [
    'bg-purple',
    'bg-info',
    'bg-pink',
    'bg-success',
    'bg-purple',
    'bg-lavandor',
    'bg-orange',
    'bg-success'
  ];

  public leaveIcon = {
    vacation: 'dashboard-heading-icon-vacation',
    unpaid: 'dashboard-heading-icon-unpaid',
    sick: 'dashboard-heading-icon-sick',
    special: 'dashboard-heading-icon-special',
    others: 'dashboard-heading-icon-others',
  };

  constructor(props) {
    super(props);

    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      },
    });

    alertify.set("notifier", "position", "top-center");


    this.state = {

      isManager: false,
      isHR: false,

      page: 0,
      rowsPerPage: 5,
      openEditPopup:false,
      openAddPopup: false,
      openDeleteConfirm: false,
      listLeaveDetails: [],
      copyListLeaveDetails: [],
      fileDetails:[],
      fileName: [],
      files:[],

      formData: {
        Id: 0,
        ApproverId: 0,
        RequesterId: 0,
        LeaveTypeId: 0,
        From: null,
        To: null,
        NoofDays: 0,
        Detail: '',
        Status: '',
        FromHalf: '1',
        ToHalf: '2',
        DocumentUrl: '',
      },

      allLeaveTypes: [],
      allWeekEndConfig: [],
      allHolidays: [],
      leaveBalance: {},
      currentUser: {},

      isview: false,
      openleavemenu: false,

      openDatePicker: false,
      strFrom: this.txtSelectDate,
      strTo: this.txtSelectDate,

      errorleavetype: null,
      errorfromto: null,

      showManager: false,
      showHR: false,
      showHolidays: false,
      showUser:true,

      disableBtn:false
    };


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
      listUrl =  "/Lists/";
    }
    

    this.init();

  }

  public init = () => {

    sp.profiles.myProperties.get().then((profile) => {
      this.setState({ isManager: profile.DirectReports.length > 0 });
    });

    this.props.graphClient
      .api("/me/manager")
      .get()
      .then((manager: any) => {
        sp.web.siteUsers.getByEmail(manager.mail).get().then((response) => {
          this.managerId = response.Id;
        });
      }).catch(error => {
        console.log(error);
      });

      // sp.web.lists.getByTitle("ConfigList").items.filter("Title eq 'Azure Group ID'").get().then((groupDet)=>{
      sp.web.getList(listUrl + "ConfigList").items.filter("Title eq 'Azure Group ID'").get().then((groupDet)=>{
        azureGroupId=groupDet[0].Value;
        sp.web.currentUser.get().then((userdata) => {
          this.props.graphClient
            .api("/groups/" + azureGroupId + "/members")
            .get()
            .then((group: any) => {
              var HRs = group.value.filter(c => c.mail == userdata.Email);
              this.setState({ isHR: HRs.length > 0 });
            }).catch(error => {
              console.log(error);
            });
    
          this.setState({ currentUser: userdata });
          this.loadAll();
        });
      });


  }

  public loadAll = () => {
    this.loadLeaveBalance();
    this.loadAppliedLeave();
    this.loadWeekEndConfig();
    this.loadHolidays()
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

  public loadAppliedLeave = () => {
    // sp.web.lists.getByTitle("LeaveRequest").items
    sp.web.getList(listUrl + "LeaveRequest").items

      .select("Id", "Title", "From", "To", "NoofDays","Created", "Detail", "Status", "LeaveType/Id", "LeaveType/Title", "LeaveType/ScreenName", "RequesterId","Requester/Title")
      .filter("RequesterId eq '" + this.state.currentUser.Id + "'")
      .orderBy('Modified', false)
      .expand("LeaveType,Requester")
      .get()
      .then((res) => {
        var lstData = this.state.listLeaveDetails;
        lstData = [];
        for (let index = 0; index < res.length; index++) {
          const leave = res[index];
          lstData.push({
            Id: leave.Id,
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
            RequesterFirstName:leave.Requester.Title
          });
        }
        this.setState({ listLeaveDetails: lstData, copyListLeaveDetails: lstData });
        this.loadLeaveTypes();
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
        for (let index = 0; index < res.length; index++) {
          const leaveType = res[index];
          allLeaveTypes.push({
            Id: leaveType.Id,
            Title: leaveType.Title,
            DisplayName: leaveType.ScreenName
          });
        }
        this.setState({ allLeaveTypes: allLeaveTypes });
      });
  }

  public loadWeekEndConfig = () => {
    // sp.web.lists.getByTitle("WeekEndConfig")
    sp.web.getList(listUrl + "WeekEndConfig")
      .items
      .filter("Holiday eq '1'")
      .get()
      .then((res) => {
        var allWeekEndConfig = this.state.allWeekEndConfig;
        allWeekEndConfig = [];
        for (let index = 0; index < res.length; index++) {
          const weekend = res[index];
          allWeekEndConfig.push({
            Id: weekend.Id,
            Title: weekend.Title
          });
        }
        this.setState({ allWeekEndConfig: allWeekEndConfig });
      });
  }

  public loadHolidays = () => {
    var currentYear = new Date().getFullYear();
    // sp.web.lists.getByTitle("Holidays")
    sp.web.getList(listUrl + "Holidays")

      .items
      .filter("Year eq '" + currentYear + "'")
      .get()
      .then((res) => {
        var allHolidays = this.state.allHolidays;
        allHolidays = [];
        for (let index = 0; index < res.length; index++) {
          const holiday = res[index];
          allHolidays.push({
            Id: holiday.Id,
            Date: holiday.Date
          });
        }
        this.setState({ allHolidays: allHolidays });
      });
  }

  public loadLeaveBalance = () => {
    var currentYear = new Date().getFullYear();
    // sp.web.lists.getByTitle("LeaveBalance")
    sp.web.getList(listUrl + "LeaveBalance")

      .items
      .filter("Year eq '" + currentYear + "' and EmployeeEmailId eq '" + this.state.currentUser.Id + "'")
      .get()
      .then((res) => {
        var leaveBalance = this.state.leaveBalance;
        if (res.length > 0) {
          leaveBalance = res[0];
        }
        this.setState({ leaveBalance: leaveBalance });
      });
  }

  public openPopup = () => {
    this.resetForm();
    this.setState({ openAddPopup: true,openEditPopup:false,disableBtn:false });
  }

  public openManager = (value) => {
    this.setState({ showManager: value, showHR: false, showHolidays: false });
    value==false?this.setState({showUser:true,showManager: value, showHR: false, showHolidays: false }):this.setState({showUser:false})
    this.loadAll();
  }

  public openHolidays = (value) => {
    this.setState({ showHolidays: value, showManager: false, showHR: false,showUser:false });
  }

  public openHR = (value) => {
    this.setState({ showHR: value, showManager: false, showHolidays: false,showUser:false });
    this.loadAll();
  }


  public closePopup = () => {
    this.setState({ openAddPopup: false,fileDetails:[],openEditPopup:false ,disableBtn:false});
  }

  public removeDoc = (e) => {
    var targetelement=e.currentTarget.className;
    var filesArray=this.state.fileDetails;
    filesArray=filesArray.filter((key,index)=>{return index!=targetelement;});
    // this.file = null;
    this.setState({ fileDetails: filesArray });
  }

  public closeViewPopup = () => {
    this.setState({ isview: false ,fileDetails:[]});
  }

  public resetForm = () => {
    var formData = this.state.formData;
    formData = {
      Id: 0,
      ApproverId: 0,
      RequesterId: 0,
      LeaveTypeId: 0,
      From: null,
      To: null,
      NoofDays: 0,
      Detail: '',
      Status: '',
      FromHalf: '1',
      ToHalf: '2',
      DocumentUrl: ''
    };
    this.file = null;
    this.setState({ formData: formData, strFrom: this.txtSelectDate, strTo: this.txtSelectDate, fileName:[] });
  }

  public setFormHalf = (value) => {
    var formData = this.state.formData;
    formData.FromHalf = value;
    this.setState({ formData: formData });
    this.calculateNoOfDays();
  }

  public setToHalf = (value) => {
    var formData = this.state.formData;
    formData.ToHalf = value;
    this.setState({ formData: formData });
    this.calculateNoOfDays();
  }

  public setLeaveType = (event: React.ChangeEvent<any>) => {
    var formData = this.state.formData;
    formData.LeaveTypeId = parseInt(event.target.value);
    this.setState({ formData: formData });
    if (!formData.LeaveTypeId) {
      this.setState({ errorleavetype: 'Leave type is required' });
    } else {
      this.setState({ errorleavetype: null });
    }
  }

  public inputChangeHandler = (e) => {
    let formData = this.state.formData;
    formData[e.target.name] = e.target.value;
    this.setState({
      formData
    });
  }

  public searchLeave = (e) => {
    var lstData = this.state.listLeaveDetails;
    var text = e.target.value;
    if (text) {
      lstData = this.state.copyListLeaveDetails.filter(c => c.Detail.toLowerCase().indexOf(text.toLowerCase()) > -1);
    } else {
      lstData = this.state.copyListLeaveDetails;
    }
    this.setState({ listLeaveDetails: lstData });
  }

  public dateChangeHandler = (e) => {
    let formData = this.state.formData;
    formData[e.target.name] = new Date(e.target.value);
    this.setState({
      formData
    });
    this.calculateNoOfDays();
  }

  public checkIfHoliday = (value: Date) => {
    for (let index = 0; index < this.state.allHolidays.length; index++) {
      const holiday = this.state.allHolidays[index];
      var date = new Date(holiday.Date).toDateString();
      if (date == value.toDateString()) {
        return true;
      }
    }
    return false;
  }

  public checkIfWeekEnd = (value: Date) => {
    var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    var dayName = days[value.getDay()];
    for (let index = 0; index < this.state.allWeekEndConfig.length; index++) {
      const weekend = this.state.allWeekEndConfig[index];
      if (weekend.Title == dayName) {
        return true;
      }
    }
    return false;
  }

  public calculateNoOfDays = () => {
    var formData = this.state.formData;
    if (formData.From > formData.To) {
      formData.NoofDays = 0;
      this.setState({ formData: formData });
      return;
    }
    formData.NoofDays = 0;
    var startDate = new Date(formData.From.toISOString());
    var endDate = new Date(formData.To.toISOString());
    var isholiday = this.checkIfHoliday(startDate);
    var isweekend = this.checkIfWeekEnd(startDate);
    if (!isholiday && !isweekend) {
      formData.NoofDays = 1;
      if (formData.FromHalf == '2') {
        formData.NoofDays = 0.5;
      }
      if ((startDate.toISOString() == endDate.toISOString()) && formData.ToHalf == '1') {
        formData.NoofDays = 0.5;
      }
    }
    startDate.setDate(startDate.getDate() + 1);
    while (startDate <= endDate) {
      var isholiday = this.checkIfHoliday(startDate);
      var isweekend = this.checkIfWeekEnd(startDate);
      if (!isholiday && !isweekend) {
        if (startDate.toISOString() == endDate.toISOString()) {
          if (formData.ToHalf == '1') {
            formData.NoofDays = formData.NoofDays + 0.5;
          } else {
            formData.NoofDays = formData.NoofDays + 1;
          }
        } else {
          formData.NoofDays = formData.NoofDays + 1;
        }
      } else if (startDate.toISOString() == endDate.toISOString()) {
        if (formData.ToHalf == '1') {
          formData.NoofDays = formData.NoofDays - 0.5;
        }
      }
      startDate.setDate(startDate.getDate() + 1);
    }
    this.setState({ formData: formData });
  }

  public editLeave = (id) => {
    this.getLeaveData(id, false);
  }

  public viewLeave = (id) => {
    this.getLeaveData(id, true);
  }

  public getLeaveData = (id, view) => {
    // sp.web.lists.getByTitle("LeaveRequest")
    sp.web.getList(listUrl + "LeaveRequest")

      .items.getById(id).get()
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

            // if(docResponse.length<=i+1)
            // this.setState( {fileDetails:docResponseFiles});
          }

        }

        if (view) {
          this.setState({ formData: formData, strFrom: fromDefaultDate, strTo: toDefaultDate, fileDetails: docResponseFiles, openAddPopup: false,openEditPopup:false, isview: true });
        } else {
          this.setState({ formData: formData, strFrom: fromDefaultDate, strTo: toDefaultDate, fileDetails: docResponseFiles, openAddPopup: false,openEditPopup:true, isview: false });
        }

      });
  }

  public deleteLeave = (id) => {
    this.deleteId = id;
    this.setState({ openDeleteConfirm: true });
  }

  public submit = () => {
    var formData = this.state.formData;
    var valid = true;
    if (formData.NoofDays == 0) {
      alertify.error('From and To date is required');
      return;
    }

    if (!formData.LeaveTypeId) {
      this.setState({ errorleavetype: 'Leave type is required' });
      return;
    } else {
      this.setState({ errorleavetype: null });
    }

    if (formData.From > formData.To) {
      alertify.error('From date is greater than To date');
      return;
    }
    
    var leaveBalance = this.state.leaveBalance;
    this.setState({disableBtn:true});
    //reset leave balance
    if (formData.Id) {
      var oldLeaveType = this.state.allLeaveTypes.filter(c => c.Id == this.oldLeaveTypeId)[0];
      var oldLeaveTypeUsed = oldLeaveType.Title + 'Used';
      var oldLeaveTypePendingApproval = oldLeaveType.Title + 'PendingApproval';
      if (formData.Status == 'Approved') {
        leaveBalance[oldLeaveTypeUsed] = (parseInt(leaveBalance[oldLeaveTypeUsed]) - this.oldNoofDays) + '';
      }
      if (formData.Status == 'Pending') {
        leaveBalance[oldLeaveTypePendingApproval] = (parseInt(leaveBalance[oldLeaveTypePendingApproval]) - this.oldNoofDays) + '';
      }
    }

    var selLeaveType = this.state.allLeaveTypes.filter(c => c.Id == formData.LeaveTypeId)[0];
    var selLeaveTypePendingApproval = selLeaveType.Title + 'PendingApproval';
    leaveBalance[selLeaveTypePendingApproval] = (parseInt(leaveBalance[selLeaveTypePendingApproval]) + formData.NoofDays) + '';

    formData.ApproverId = this.managerId;
    formData.RequesterId = this.state.currentUser.Id;
    formData.Status = 'Pending';

    if (formData.Id) {
      if (!this.state.fileDetails) {
        formData.DocumentUrl = '';
      }
      // sp.web.lists.getByTitle("LeaveRequest")
      sp.web.getList(listUrl + "LeaveRequest")

        .items.getById(formData.Id)
        .update(formData)
        .then((response) => {
          this.updateLeaveBalance(leaveBalance);
          // if (this.state.fileDetails) {
           this.uploadfile(formData.Id);
            alertify.success('Leave updated successfully');
          // } else {
            // alertify.success('Leave updated successfully');
          // }
        });
    } else {
      // sp.web.lists.getByTitle("LeaveRequest")
      sp.web.getList(listUrl + "LeaveRequest")

        .items.add(formData)
        .then((res) => {
          this.updateLeaveBalance(leaveBalance);
          // if (this.state.fileDetails) {
            this.uploadfile(res.data.Id);
          //   alertify.success('Leave applied successfully');
          // } else {
            alertify.success('Leave applied successfully');
          // }
        });
    }
  }

  public updateLeaveBalance = (leaveBalance) => {
    // sp.web.lists.getByTitle("LeaveBalance")
    sp.web.getList(listUrl + "LeaveBalance")

      .items.getById(leaveBalance.Id)
      .update(leaveBalance)
      .then((response) => {
        this.loadAll();
        this.closePopup();
      });
  }

  public async uploadfile (newId){
    var allUploadFiles=this.state.fileDetails;
    var siteURL = this.props.siteUrl;
    var filePath="";
   await  sp.web.getFolderByServerRelativeUrl(folderPath).folders.add(folderPath + '/' + newId).then(result => {
     if(allUploadFiles.length>0)
     {
      allUploadFiles.map((eachfileDetails,index)=>{
        if(eachfileDetails.files["name"])
        {
          result.folder.files.add(eachfileDetails.filenname, eachfileDetails.files,true)
          .then((fresult) => {
  
              var serverURL=fresult.data.ServerRelativeUrl;
              if(filePath)
               filePath = filePath+siteURL +serverURL+";";
               else
               filePath = siteURL +serverURL+";";
  
  
               if(allUploadFiles.length<=index+1)
               {
                 this.setState({ fileDetails: [] });
        
                //  sp.web.lists.getByTitle("LeaveRequest")
                sp.web.getList(listUrl + "LeaveRequest")

                 .items.getById(newId).update({ DocumentUrl: filePath }).then(function (result) {
                 });
               }

             
          });
        }
        else
       {
        if(filePath)
         filePath = filePath+eachfileDetails.files+";";
         else
         filePath = eachfileDetails.files+";";

         if(allUploadFiles.length<=index+1)
         {
           this.setState({ fileDetails: [] });
  
          //  sp.web.lists.getByTitle("LeaveRequest")
          sp.web.getList(listUrl + "LeaveRequest")

           .items.getById(newId).update({ DocumentUrl: filePath }).then(function (result) {
           });
         }
       }
       


      });
     }
     else{

        this.setState({ fileDetails: [] });

        // sp.web.lists.getByTitle("LeaveRequest")
        sp.web.getList(listUrl + "LeaveRequest")

        .items.getById(newId).update({ DocumentUrl: "" }).then(function (result) {
        });
     }


    });
  }

  public closeDelete = () => {
    this.setState({ openDeleteConfirm: false });
  }

  public confirmDelete = () => {
    var leaveData = this.state.listLeaveDetails.filter(c => c.Id == this.deleteId)[0];
    var leaveBalance = this.state.leaveBalance;
    var selLeaveType = this.state.allLeaveTypes.filter(c => c.Id == leaveData.LeaveTypeId)[0];
    var selTypeUsed = selLeaveType.Title + 'Used';
    var selTypePendingApproval = selLeaveType.Title + 'PendingApproval';

    if (leaveData.Status == 'Approved') {
      leaveBalance[selTypeUsed] = (parseInt(leaveBalance[selTypeUsed]) - leaveData.NoofDays) + '';
    } else if (leaveData.Status == 'Pending') {
      leaveBalance[selTypePendingApproval] = (parseInt(leaveBalance[selTypePendingApproval]) - leaveData.NoofDays) + '';
    }

    this.updateLeaveBalance(leaveBalance);
    // sp.web.lists.getByTitle("LeaveRequest")
    sp.web.getList(listUrl + "LeaveRequest")
      .items.getById(this.deleteId)
      .get()
      .then((response) => {
        response.Status = 'Cancelled';
        // sp.web.lists
        //   .getByTitle("LeaveRequest")
        sp.web.getList(listUrl + "LeaveRequest")
          .items.getById(response.Id)
          .update(response)
          .then((updateresponse) => {
            this.setState({ openDeleteConfirm: false });
            alertify.success('Leave cancelled successfully');
            this.loadAll();
          });
      });


  }

  public showDatePicker = (value) => {
    this.setState({ openDatePicker: value });
  }

  public setDateRange = (range) => {
    let formData = this.state.formData;
    formData.From = range.startDate;
    formData.To = range.endDate;
    this.setState({
      formData: formData, strFrom: this.formatDate(range.startDate), strTo: this.formatDate(range.endDate)
    });
    this.calculateNoOfDays();
    this.setState({ openDatePicker: false });
  }

  // public handleClose = (event) => {
  //   if (this.anchorRef.current && this.anchorRef.current.contains(event.target)) {
  //     return;
  //   }
  //   this.setState({ openleavemenu: false });
  // }

  public handleMenuItemClick = (event, leaveTypeId) => {
    this.resetForm();
    var formData = this.state.formData;
    formData.LeaveTypeId = leaveTypeId;
    this.setState({ formData: formData, openAddPopup: true, openleavemenu: false });
  }

  public fileUpload = (e) => {
    
    var files = e.target.files;
    if (files && files.length > 0) {
      var allfiles=[];
      for(let i=0;i<files.length;i++)
      {
        if(this.state.fileDetails)
        allfiles=this.state.fileDetails;
        var sepArray=allfiles.filter((eleFile)=>{return eleFile.filenname==files[i].name})
        if(sepArray.length<=0)
        {
          this.file = files[i];
          allfiles.push({filenname:files[i].name,files:this.file});
        }

        if(files.length<=i+1)
        this.setState( {fileDetails:allfiles});
        


      }

      e.target.value=null;
      
    } else {
      this.file = null;
      this.setState({ fileName: [] });
      e.target.value=null;
    }
  }

  public render(): React.ReactElement<IUngotiApplyLeaveProps> {

    // const anchorRef = React.useRef(null);

    const columns = [
      { field: 'LeaveType', title: 'Type' },
      { field: 'RequestedDate', title: 'Requested Date' },
      { field: 'strFrom', title: 'From' },
      { field: 'strTo', title: 'To' },
      { field: 'strNoofDays', title: 'No. of days' },
      { field: 'Status', title: 'Status' },
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


    const handleChangePage = (event: unknown, newPage: number) => {
      this.setState({ page: newPage });
    };

    const handleChangeRowsPerPage = (event: React.ChangeEvent<HTMLInputElement>) => {
      // this.setState({ rowsPerPage: this.state.rowsPerPage + parseInt(event.target.value) });
      this.setState({ rowsPerPage: parseInt(event.target.value, 10) });
      this.setState({ page: 0 });
    };

    var propscardcolor =this.props.color;

    var commonBackgroundColor={
      backgroundColor:propscardcolor
    };
    

    return (


      <div className={styles.ungotiApplyLeave}>

        <section className="page-section">
          <div className="page-title">
            <Grid container spacing={2} justify="space-between" >
              <Typography component={'h3'}>
                {
                 // this.props.card && !this.state.showHolidays ? this.props.cardTitle : '' 
                 this.state.showUser&&this.props.card?this.props.cardTitle:this.state.showManager&&this.props.chkmanager?this.props.managerTitle:this.state.showHR&&this.props.chkHR?this.props.hrTitle:this.state.showHolidays&&this.props.chkHolidays?this.props.holidaysTitle:""
                }
              </Typography>
              <ButtonGroup disableElevation variant="contained" size="small" color="primary" className="role-button-group">
                {
                  this.state.isManager&&this.props.chkmanager ? <Button  size="small"  className={this.state.showManager?'active':''} onClick={this.openManager.bind(this, true)}>Manager</Button> : ''
                }

                {
                  this.state.isHR&&this.props.chkHR ? <Button size="small"    className={this.state.showHR?'active':''} onClick={this.openHR.bind(this, true)}>HR</Button> : ''
                }
                {
                this.props.UserRequest ?<Button size="small"   className={this.state.showUser?'active':''}  onClick={this.openManager.bind(this, false)}>User</Button>:''

                }
                {
                  this.props.chkHolidays?<Button size="small"   className={this.state.showHolidays?'active':''} onClick={this.openHolidays.bind(this, true)}>Holidays</Button>
                  :''
                }
                {
                  this.state.showManager||this.state.showHR||this.state.showHolidays||!this.props.UserRequest?'':<Button className="radius-button" color="secondary" variant="contained" size="small" onClick={this.openPopup}>New request</Button>
                }

              </ButtonGroup>


            </Grid>
          </div>
        </section>

        {
          (!this.state.showManager && !this.state.showHR && !this.state.showHolidays && this.props.UserRequest) ?
            <div>

              <div>

                <section className="page-section">
                  <Grid container spacing={2}>

                    {
                      this.props.card ?
                        <Grid item xs={12} sm={12} md={12} lg={12} xl={12}>
                          <Grid container spacing={2}>

                            {
                              this.state.allLeaveTypes.map((leaveType, index) => {

                                var totalLeave = this.state.leaveBalance[leaveType.Title];
                                var usedLeave = this.state.leaveBalance[leaveType.Title + 'Used'];
                                var pendingLeave = this.state.leaveBalance[leaveType.Title + 'PendingApproval'];

                                var available = parseFloat(totalLeave) - parseFloat(usedLeave);

                               // var cardcolor = "dashboard-card " + this.leaveColors[index];
                              var cardcolor = "dashboard-card ";

                                var propscardcolor =this.props.color;

                                var cardLeftBorderStyle = {
                                   borderLeftColor:  this.props.color
                                };
                            //     var iconcardBorderStyle = {
                            //       // -webkit-box-shadow: 0 0 5px #f9c91f!important;
                            //       boxShadow: "0 0 5px"+ propscardcolor,
                            //       border: "1px solid"+ propscardcolor
                            //    };
                            //    var cardHeaderStyle = {
                            //     // -webkit-box-shadow: 0 0 5px #f9c91f!important;
                            //     color: propscardcolor
                            //  };

                                var progressValue = (parseFloat(usedLeave) / parseFloat(totalLeave)) * 100;

                                if (parseFloat(usedLeave) > parseFloat(totalLeave)) {
                                  progressValue = 100;
                                }
                                if (available < 0) {
                                  available = 0;
                                }

                                var leaveIcon = this.leaveIcon.others;
                                if (leaveType.DisplayName.toLowerCase().indexOf('vacation') >= 0) {
                                  leaveIcon = this.leaveIcon.vacation;
                                } else if (leaveType.DisplayName.toLowerCase().indexOf('sick') >= 0) {
                                  leaveIcon = this.leaveIcon.sick;
                                } else if (leaveType.DisplayName.toLowerCase().indexOf('special') >= 0) {
                                  leaveIcon = this.leaveIcon.special;
                                } else if (leaveType.DisplayName.toLowerCase().indexOf('paid') >= 0) {
                                  leaveIcon = this.leaveIcon.unpaid;
                                }

                                return (
                                  <Grid className="dashboard-grid" item xs={12} sm={4} md={2} lg={"auto"} xl={"auto"}>
                                    <Paper elevation={2} square={false} style={cardLeftBorderStyle} className={cardcolor}>
                                      <div className="heading-group">
                                        <div  className={'dashboard-heading-icon ' + leaveIcon}>
                                        </div>
                                        <div className={'dashboard-heading'} >
                                          <Typography component={'h6'}  >
                                            {leaveType.DisplayName}
                                          </Typography>
                                          <Typography component={'h2'}  className={"card-totalnumber"}>
                                            {totalLeave}
                                          </Typography>
                                        </div>
                                      </div>
                                      {/* <div className="dashboard-chart">
                                        <LinearProgress variant="determinate" className="dashboard-chart-progress" value={progressValue} />

                                      </div> */}
                                      <div className="dashboard-group">
                                        <List className="dashboard-list" >
                                          <ListItem>
                                            <ListItemText>
                                              Available <Badge>{available}</Badge>
                                            </ListItemText>

                                          </ListItem>
                                          <ListItem>
                                            <ListItemText>
                                              Consumed <Badge>{usedLeave}</Badge>
                                            </ListItemText>

                                          </ListItem>
                                          <ListItem>
                                            <ListItemText>
                                              Pending <Badge>{pendingLeave}</Badge>
                                            </ListItemText>

                                          </ListItem>

                                        </List>

                                      </div>

                                    </Paper>
                                  </Grid>
                                );
                              })
                            }

                          </Grid>
                        </Grid>
                        : ''
                    }


                    {
                      this.props.list ?
                        <Grid className="manageLeave" item xs={12} sm={12} md={12} lg={12} xl={12}>

                          <MaterialTable
                            title={this.props.listTitle}
                            icons={tableIcons}
                            columns={columns}
                            data={this.state.listLeaveDetails}
                            actions={[
                              (rowData: LeaveDetails) => ({
                                icon: forwardRef((props: any, ref: any) => <EditIcon />),
                                tooltip: 'Edit',
                                onClick: (event, value) => this.editLeave(rowData.Id),
                                disabled: (rowData["Status"] == 'Cancelled'  || rowData.From <= currentDate||rowData["Status"] == 'Rejected')
                              }),
                              (rowData: LeaveDetails) => ({
                                icon: forwardRef((props: any, ref: any) => <DeleteIcon />),
                                tooltip: 'Cancel',
                                onClick: (event, value) => this.deleteLeave(rowData.Id),
                                disabled: ((rowData["Status"] == 'Rejected') || (rowData["Status"] == 'Cancelled') || rowData.From <= currentDate)
                              }),
                              {
                                icon: forwardRef((props: any, ref: any) => <VisibilityIcon />),
                                tooltip: 'View',
                                onClick: (event, rowData: LeaveDetails) => this.viewLeave(rowData.Id),
                              }
                            ]}
                            options={{
                              actionsColumnIndex: 6
                            }}
                          />

                        </Grid>
                        : ''
                    }


                  </Grid>
                </section>
              </div>

              <Dialog open={this.state.openAddPopup||this.state.openEditPopup} className="applyLeaveDialog">
                <DialogTitle className="MuiDialogTitle-bg" id="form-dialog-title">
                  <Typography component={"h5"}>Leave Request for {this.state.currentUser.Title}</Typography>
                </DialogTitle>
                <DialogContent>
                  <section className="dateRangePicker">


                    <Grid container spacing={2} className="datefield">
                      <Grid sm={12} md={6} onClick={this.showDatePicker.bind(this, true)}>
                        <Typography component={"p"} className="small-text">
                          FROM DATE
               </Typography>

                        <Typography component={"p"}>
                          {this.state.strFrom}
                        </Typography>

                      </Grid>
                      <div className="dateRangePicker-totalDays" onClick={this.showDatePicker.bind(this, true)}>
                        <span className="number">{this.state.formData.NoofDays} Day(s)</span>

                      </div>
                      <Grid sm={12} md={6} className={"text-right"} onClick={this.showDatePicker.bind(this, true)}>
                        <Typography component={"p"} className="small-text">
                          TO DATE
               </Typography>

                        <Typography component={"p"}>
                          {this.state.strTo}
                        </Typography>

                      </Grid>

                      <DateRangePicker
                        // definedRanges={[]}
                        open={this.state.openDatePicker}
                        toggle={this.showDatePicker.bind(this, false)}
                        onChange={(range) => this.setDateRange(range)}
                        wrapperClassName="modal-DateRagePicker"
                      />

                    </Grid>


                  </section>

                  <Grid container justify={"space-between"} alignItems={"center"}>
                    <Grid sm={5} className="text-right">
                      <ButtonGroup color="primary" variant="contained" size={"small"} disableElevation>
                        <Button variant={"contained"} color={this.state.formData.FromHalf == '1' ? 'primary' : 'default'} onClick={this.setFormHalf.bind(this, '1')}>First half</Button>
                        <Button variant={"contained"} color={this.state.formData.FromHalf == '2' ? 'primary' : 'default'} onClick={this.setFormHalf.bind(this, '2')}>Second half</Button>
                      </ButtonGroup>
                    </Grid>
                    <Grid sm={2}>
                      <Typography component={"p"} className={"small text-center"}>To</Typography>
                    </Grid>
                    <Grid sm={5}>
                      <ButtonGroup color="primary" variant="contained" size={"small"} disableElevation>
                        <Button variant={"contained"} color={this.state.formData.ToHalf == '1' ? 'primary' : 'default'} onClick={this.setToHalf.bind(this, '1')}>First half</Button>
                        <Button variant={"contained"} color={this.state.formData.ToHalf == '2' ? 'primary' : 'default'} onClick={this.setToHalf.bind(this, '2')}>Second half</Button>
                      </ButtonGroup>
                    </Grid>
                  </Grid>
                  <Grid container>
                    <Grid sm={12} >
                      <FormControl variant="outlined" className="form-group" size="small" error={this.state.errorleavetype ? true : false}>
                        <InputLabel id="standard-select-currency" >Leave Type</InputLabel>
                        <Select
                          labelId="standard-select-currency"
                          id="standard-select-currency"
                          value={this.state.formData.LeaveTypeId} onChange={this.setLeaveType}
                          label="Leave Type"
                        >
                          {
                            this.state.allLeaveTypes.map((leaveType) => {
                              return (
                                <MenuItem value={leaveType.Id}>{leaveType.DisplayName}</MenuItem>
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
                        value={this.state.formData.Detail}
                        onChange={(e) => this.inputChangeHandler.call(this, e)}
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


                  <Grid container spacing={2}>
                    <Grid item xs={12} sm={12} md={12} lg={12} xl={12} className="form-group">
                      {/* <input accept="image/*" type="file" id="icon-button-file" multiple onChange={(e) => this.fileUpload.call(this, e)} style={{ visibility: "hidden" }} />
                      <label htmlFor="icon-button-file" className="uploadbtn">
                        <Button color="primary" aria-label="upload picture" component="span" variant="contained" >
                          <AttachmentIcon />Add Documents
                        </Button>
                        </label> */}
                      
                        {
                          this.state.fileDetails?
                          this.state.openAddPopup?
                            this.state.fileDetails.map((filedet,index) => {
                           
                              return (
                                <>
                                <label>{filedet.filenname}</label>
                                <IconButton color="secondary">
                                {
                                   filedet.filenname ? <span  onClick={this.removeDoc.bind(this)} className={index.toString()}><DeleteForeverIcon /></span> : ''
                                }
                              </IconButton><br></br>
                              </>
                              );
                            
                            }):  this.state.fileDetails.map((filedet,index) => {
                           
                              return (
                                <>
                                <label><a href={filedet.files} target="_blank">{filedet.filenname}</a></label>
                                <IconButton color="secondary">
                                {
                                   filedet.filenname ? <span  onClick={this.removeDoc.bind(this)} className={index.toString()}><DeleteForeverIcon /></span> : ''
                                }
                              </IconButton><br></br>
                              </>
                              );
                            
                            }):""
                          } 
                          <input  type="file" id="icon-button-file" multiple onChange={(e) => this.fileUpload.call(this, e)} style={{ visibility: "hidden" }} />
                      <label htmlFor="icon-button-file" className="uploadbtn">
                        <Button size="small" color="primary" aria-label="upload picture" component="span" variant="contained" >
                          <AttachmentIcon />Add Documents
                        </Button>
                        </label>
                      
                      {/* {
                            this.state.fileDetails.map((filedet) => {
                              return (
                                <IconButton color="secondary">
                                {
                                  filedet.filenname ? <DeleteForeverIcon onClick={this.removeDoc} /> : ''
                                }
                              </IconButton>
                              );
                            })
                          }  */}

                    </Grid>
                    {/* <Grid item xs={12} sm={12} md={6} lg={6} xl={6} className="form-group">
                     
                    </Grid> */}
                  </Grid>

                </DialogContent> 
                <DialogActions>
                  <Button variant="contained" disableElevation color="default" size="small" onClick={this.closePopup}>
                    Cancel
          </Button> 
                  <Button variant="contained" disabled={this.state.disableBtn} disableElevation color="primary" size="small" onClick={this.submit}>
                    Apply
          </Button>

                </DialogActions>
              </Dialog>




              <Dialog open={this.state.isview} className="applyLeaveDialog">
                <DialogTitle className="MuiDialogTitle-bg" id="form-dialog-title">
                  <Typography component={"h5"}>Leave Details for {this.state.currentUser.Title}</Typography>
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





              <Dialog
                open={this.state.openDeleteConfirm}
                onClose={this.closeDelete}
                aria-labelledby="alert-dialog-title"
                aria-describedby="alert-dialog-description"
              >
                <DialogTitle id="alert-dialog-title">{"Leave Cancellation?"}</DialogTitle>
                <DialogContent>
                  <DialogContentText id="alert-dialog-description">
                    Do you want to cancel the leave?
          </DialogContentText>
                </DialogContent>
                <DialogActions>
                  <Button disableElevation variant="contained" onClick={this.closeDelete} color="default">
                    No
          </Button>
                  <Button disableElevation variant="contained" onClick={this.confirmDelete} color="primary" autoFocus>
                    Yes
          </Button>
                </DialogActions>
              </Dialog>
            </div>
            :

            ''
        }

        {
          this.state.showManager && this.props.chkmanager ?
            <div>
              {/* <section className="page-section">
                <div className="page-title">
                  <Grid container spacing={2} justify="space-between" >
                    <Typography component={'h3'}>
                      {
                        this.props.card ? this.props.cardTitle : ''
                      }
                    </Typography>
                    <ButtonGroup disableElevation variant="contained" size="small" color="primary">
                      <Button size="small" onClick={this.openManager.bind(this, false)}>User</Button>
                    </ButtonGroup>

                  </Grid>
                </div>
              </section> */}
              <Manager  currentContext={this.props.currentContext} />

            </div>
            : ''
        }

        {
          this.state.showHolidays && this.props.chkHolidays && <Holidays  currentContext={this.props.currentContext} />
        }

        {
          this.state.showHR && this.props.chkHR ?
            <div>
              {/* <section className="page-section">
                <div className="page-title">
                  <Grid container spacing={2} justify="space-between" >
                    <Typography component={'h3'}>
                      {
                        this.props.card ? this.props.cardTitle : ''
                      }
                    </Typography>
                    <ButtonGroup disableElevation variant="contained" size="small" color="primary">
                      <Button size="small" onClick={this.openManager.bind(this, false)}>User</Button>
                    </ButtonGroup>

                  </Grid>
                </div>
              </section> */}

              <HR graphClient={this.props.graphClient} currentContext={this.props.currentContext} />

            </div>
            : ''
        }

      </div >
    );
  }
}
