import * as React from 'react';
import { makeStyles } from '@material-ui/core/styles';
import List from '@material-ui/core/List';
import ListItem from '@material-ui/core/ListItem';
import ListItemText from '@material-ui/core/ListItemText';
import ListItemAvatar from '@material-ui/core/ListItemAvatar';
import Avatar from '@material-ui/core/Avatar';
import ImageIcon from '@material-ui/icons/Image';
import WorkIcon from '@material-ui/icons/Work';
import BeachAccessIcon from '@material-ui/icons/BeachAccess';

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


import '../../../scss/styles.scss';
import "alertifyjs";

import "../../../ExternalRef/CSS/style.css";
import "../../../ExternalRef/CSS/alertify.min.css";
import { Grid, Card } from '@material-ui/core';
var alertify: any = require("../../../ExternalRef/JS/alertify.min.js");
export interface IHolidayProps {
  currentContext: any;
}
var listUrl="";
export default class Holidays extends React.Component<IHolidayProps, any> {

  constructor(props) {
    super(props);
    this.state = {
      allHolidays: [],
    };
    alertify.set("notifier", "position", "top-right");

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

    this.loadHolidays();
  }

  public loadHolidays = () => {
    var year = new Date().getFullYear();
    let timeNow = new Date().toISOString();
    // sp.web.lists.getByTitle("Holidays")
    
    
    sp.web.getList(listUrl + "Holidays")
      .items
      .filter("Date ge '" + timeNow + "'")
      .orderBy('Date')
      .get()
      .then((res) => {
        var holidays = [];
        for (let index = 0; index < res.length; index++) {
          const element = res[index];
          holidays.push({
            Title: element.Title,
            Date:element.Date
          });
        }
        this.setState({ allHolidays: holidays });
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

  public render(): React.ReactElement {
    return (
      <Card elevation={0}>
        <List>
          <Grid container>
           
              {
                this.state.allHolidays.map((holiday) => {
                  var loopmonth=new Date(holiday.Date).getMonth();
                  var loopDate=new Date(holiday.Date).getDate();
                  var fullDate=this.formatDate(new Date(holiday.Date));
                  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                  var monthName=monthNames[loopmonth];
                  return (
                    <Grid item sm={3}>
                    <ListItem className="list-item">
                      <div className="calendar-container bg-purple">
                  <div className="calendar-header">{monthName}</div>
                        <div className="calendar-body">{loopDate}</div>
                      </div>
                      <ListItemText primary={holiday.Title} secondary={fullDate} />
                    </ListItem>
                    </Grid>
                  );
                })
              }
            
          </Grid>
        </List>
      </Card>
    );
  }
}
