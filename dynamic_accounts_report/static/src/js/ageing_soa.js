odoo.define('dynamic_accounts_report.ageing_soa', function (require) {
    'use strict';
    var AbstractAction = require('web.AbstractAction');
    var core = require('web.core');
    var field_utils = require('web.field_utils');
    var rpc = require('web.rpc');
    var session = require('web.session');
    var utils = require('web.utils');
    var QWeb = core.qweb;
    var round_di = utils.round_decimals;
    var _t = core._t;

    window.click_num = 0;
    var PartnerAgeingSOA = AbstractAction.extend({
    template: 'AgeingTempSOA',
        events: {
            'click .parent-line': 'journal_line_click',
            'click .child_col1': 'journal_line_click',
            'click #apply_filter': 'apply_filter',
            'click #pdf': 'print_pdf',
            'click #xlsx': 'print_xlsx',
            'click .gl-line': 'show_drop_down',
            'click .view-account-move': 'view_acc_move',
        },

        init: function(parent, action) {
        this._super(parent, action);
                this.currency=action.currency;
                this.report_lines = action.report_lines;

                this.wizard_id = action.context.wizard | null;

            },


          start: function() {
            var self = this;
            self.initial_render = true;
            rpc.query({
                model: 'account.partner.ageing.soa',
                method: 'create',
                args: [{

                }]
            }).then(function(t_res) {
                self.wizard_id = t_res;

                self.load_data(self.initial_render);
            })
        },


        load_data: function (initial_render = true) {

            var self = this;

                self.$(".categ").empty();
                try{
                    var self = this;
                    self._rpc({
                        model: 'account.partner.ageing.soa',
                        method: 'view_report',
                        args: [[this.wizard_id]],
                    }).then(function(datas) {
                    _.each(datas['report_lines'][0], function(rep_lines) {
                            rep_lines.total = self.format_currency(datas['currency'],rep_lines.total);
                            rep_lines[6]['amount'] = self.format_currency(datas['currency'],rep_lines[6]['amount']);
                            rep_lines[5]['amount'] = self.format_currency(datas['currency'],rep_lines[5]['amount']);
                            rep_lines[4]['amount'] = self.format_currency(datas['currency'],rep_lines[4]['amount']);
                            rep_lines[3]['amount'] = self.format_currency(datas['currency'],rep_lines[3]['amount']);
                            rep_lines[2]['amount'] = self.format_currency(datas['currency'],rep_lines[2]['amount']);
                            rep_lines[1]['amount'] = self.format_currency(datas['currency'],rep_lines[1]['amount']);
                            rep_lines[0]['amount'] = self.format_currency(datas['currency'],rep_lines[0]['amount']);
                            
                            
                            rep_lines[6]['duty_amount'] = self.format_currency(datas['currency'],rep_lines[6]['duty_amount']);
                            rep_lines[5]['duty_amount'] = self.format_currency(datas['currency'],rep_lines[5]['duty_amount']);
                            rep_lines[4]['duty_amount'] = self.format_currency(datas['currency'],rep_lines[4]['duty_amount']);
                            rep_lines[3]['duty_amount'] = self.format_currency(datas['currency'],rep_lines[3]['duty_amount']);
                            rep_lines[2]['duty_amount'] = self.format_currency(datas['currency'],rep_lines[2]['duty_amount']);
                            rep_lines[1]['duty_amount'] = self.format_currency(datas['currency'],rep_lines[1]['duty_amount']);
                            rep_lines[0]['duty_amount'] = self.format_currency(datas['currency'],rep_lines[0]['duty_amount']);
                            
                            rep_lines[6]['vat_amount'] = self.format_currency(datas['currency'],rep_lines[6]['vat_amount']);
                            rep_lines[5]['vat_amount'] = self.format_currency(datas['currency'],rep_lines[5]['vat_amount']);
                            rep_lines[4]['vat_amount'] = self.format_currency(datas['currency'],rep_lines[4]['vat_amount']);
                            rep_lines[3]['vat_amount'] = self.format_currency(datas['currency'],rep_lines[3]['vat_amount']);
                            rep_lines[2]['vat_amount'] = self.format_currency(datas['currency'],rep_lines[2]['vat_amount']);
                            rep_lines[1]['vat_amount'] = self.format_currency(datas['currency'],rep_lines[1]['vat_amount']);
                            rep_lines[0]['vat_amount'] = self.format_currency(datas['currency'],rep_lines[0]['vat_amount']);

                            rep_lines['direction'] = self.format_currency(datas['currency'],rep_lines['direction']);
                            rep_lines['unalloc'] = self.format_currency(datas['currency'],rep_lines['unalloc']);

                             });

                            if (initial_render) {
                                    self.$('.filter_view_tb').html(QWeb.render('AgeingFilterViewSOA', {
                                        filter_data: datas['filters'],
                                    }));
                                    self._rpc({
                                                model: 'account.partner.ageing.soa',
                                                method: 'get_partners',
                                                //args: [[this.wizard_id]],
                                            }).then(function(datas) {
                                                self.$el.find('.partners').select2({
                                                    placeholder: 'Partners...',
                                                    multiple: true,
                                                    data: datas,
                                                    query: function(q) {
                                                              var pageSize,
                                                                results,
                                                                that = this;
                                                              pageSize = 20; // or whatever pagesize
                                                              results = [];
                                                              if (q.term && q.term !== '') {
                                                                // HEADS UP; for the _.filter function i use underscore (actually lo-dash) here
                                                                results = _.filter(that.data, function(e) {
                                                                  return e.text.toUpperCase().indexOf(q.term.toUpperCase()) >= 0;
                                                                });
                                                              } else if (q.term === '') {
                                                                results = that.data;
                                                              }
                                                              q.callback({
                                                                results: results.slice((q.page - 1) * pageSize, q.page * pageSize),
                                                                more: results.length >= q.page * pageSize,
                                                              });
                                                            },
                                                    /*query: function(q) {
                                                        alert(q);
                                                        },*/
                                            });
                                            });
                                    self.$el.find('.category').select2({
                                        placeholder: ' Partner Category...',
                                    });
                                    self.$el.find('.target_move').select2({
                                        placeholder: ' Target Move...',
                                    });
                                    self.$el.find('.result_selection').select2({
                                        placeholder: ' Account Type...',
                                    });

                            }
                            var child=[];

                        self.$('.table_view_tb').html(QWeb.render('AgeingtableSOA', {

                                            report_lines : datas['report_lines'],
                                            move_lines :datas['report_lines'][2],
                                            filter : datas['filters'],
                                            currency : datas['currency'],
                                        }));

                });

                    }
                catch (el) {
                    window.location.href
                    }
            },
            format_currency: function(currency, amount) {
            return this.format_currency_no_symbol(amount, currency[3], currency);
    
            /*if (currency[1] === 'after') {
                return amount + ' ' + (currency[0] || '');
            } else {
                return (currency[0] || '') + ' ' + amount;
            }*/
        },
    
        format_currency_no_symbol: function(amount, decimals, currency) {
            if (typeof amount === 'number') {
                amount = round_di(amount, decimals).toFixed(decimals);
                amount = field_utils.format.float(round_di(amount, decimals), {
                    digits: [69, decimals],
                });
            }
    
            return amount;
        },

            print_pdf: function(e) {
            e.preventDefault();

            var self = this;
            self._rpc({
                model: 'account.partner.ageing.soa',
                method: 'view_report',
                args: [
                    [self.wizard_id]
                ],
            }).then(function(data) {
                var action = {
                    'type': 'ir.actions.report',
                    'report_type': 'qweb-pdf',
                    'report_name': 'dynamic_accounts_report.partner_ageing_soa',
                    'report_file': 'dynamic_accounts_report.partner_ageing_soa',
                    'data': {
                        'report_data': data
                    },
                    'context': {
                        'active_model': 'account.partner.ageing.soa',
                        'landscape': 1,
                        'ageing_pdf_report': true

                    },
                    'display_name': 'Duty Partner Ageing',
                };

                return self.do_action(action);
            });
        },



        print_xlsx: function() {
            var self = this;
            self._rpc({
                model: 'account.partner.ageing.soa',
                method: 'view_report',
                args: [
                    [self.wizard_id]
                ],
            }).then(function(data) {

                var action = {
                    'type': 'ir_actions_dynamic_xlsx_download',
                    'data': {
                         'model': 'account.partner.ageing.soa',
                         'options': JSON.stringify(data['filters']),
                         'output_format': 'xlsx',
                         'report_data': JSON.stringify(data['report_lines']),
                         'report_name': 'Duty Partner Ageing',
                         'dfr_data': JSON.stringify(data),
                    },
                };
                return self.do_action(action);
            });
        },





        create_lines_with_style: function(rec, attr, datas) {
            var temp_str = "";
            var style_name = "border-bottom: 1px solid #e6e6e6;";
            var attr_name = attr + " style="+style_name;



            temp_str += "<td  class='child_col1' "+attr_name+" >"+rec['code'] +rec['name'] +"</td>";
            if(datas.currency[1]=='after'){
            temp_str += "<td  class='child_col2' "+attr_name+" >"+rec['debit'].toFixed(2)+datas.currency[0]+"</td>";
            temp_str += "<td  class='child_col3' "+attr_name+" >"+rec['credit'].toFixed(2) +datas.currency[0]+ "</td>";

            }
            else{
            temp_str += "<td  class='child_col2' "+attr_name+" >"+datas.currency[0]+rec['debit'].toFixed(2) + "</td>";
            temp_str += "<td  class='child_col3' "+attr_name+">"+datas.currency[0]+rec['credit'].toFixed(2) + "</td>";

            }
            return temp_str;
        },


        journal_line_click: function (el){

            click_num++;
            var self = this;
            var line = $(el.target).parent().data('id');

            return self.do_action({
                type: 'ir.actions.act_window',
                    view_type: 'form',
                    view_mode: 'form',
                    res_model: 'account.move',
                    views: [
                        [false, 'form']
                    ],
                    res_id: line,
                    target: 'current',
            });

        },

        show_drop_down: function(event) {
            event.preventDefault();
            var self = this;
            var account_id = $(event.currentTarget).data('account-id');

            var partner_id = $(event.currentTarget)[0].cells[0].innerText;

            var offset = 0;
            var td = $(event.currentTarget).next('tr').find('td');
            if (td.length == 1) {

                    self._rpc({
                        model: 'account.partner.ageing.soa',
                        method: 'view_report',
                        args: [
                            [self.wizard_id]
                        ],
                    }).then(function(data) {


                    _.each(data['report_lines'][0], function(rep_lines) {
                    _.each(rep_lines['child_lines'], function(child_line) {
                            child_line.amount = self.format_currency(data['currency'],child_line.amount);
                            child_line.paid_amount = self.format_currency(data['currency'],child_line.paid_amount);

                            });
                             });


                    for (var i = 0; i < data['report_lines'][0].length; i++) {
                    if (account_id == data['report_lines'][0][i]['partner_id'] ){
                    $(event.currentTarget).next('tr').find('td .gl-table-div').remove();
                    $(event.currentTarget).next('tr').find('td ul').after(
                        QWeb.render('SubSectionalSOA', {
                            account_data: data['report_lines'][0][i]['child_lines'],
                        }))
                    $(event.currentTarget).next('tr').find('td ul li:first a').css({
                        'background-color': '#00ede8',
                        'font-weight': 'bold',
                    });
                     }
                    }

                    });
            }
        },

        view_acc_move: function(event) {
            event.preventDefault();
            var self = this;
            var context = {};
            var show_acc_move = function(res_model, res_id, view_id) {
                var action = {
                    type: 'ir.actions.act_window',
                    view_type: 'form',
                    view_mode: 'form',
                    res_model: res_model,
                    views: [
                        [view_id || false, 'form']
                    ],
                    res_id: res_id,
                    target: 'current',
                    context: context,
                };
                return self.do_action(action);
            };
            rpc.query({
                    model: 'account.move',
                    method: 'search_read',
                    domain: [
                        ['id', '=', $(event.currentTarget).data('move-id')]
                    ],
                    fields: ['id'],
                    limit: 1,
                })
                .then(function(record) {
                    if (record.length > 0) {
                        show_acc_move('account.move', record[0].id);
                    } else {
                        show_acc_move('account.move', $(event.currentTarget).data('move-id'));
                    }
                });
        },


        apply_filter: function(event) {

            event.preventDefault();
            var self = this;
            self.initial_render = false;

            var filter_data_selected = {};

            if ($("#date_from").val()) {
                var dateString = $("#date_from").val();

                filter_data_selected.date_from= dateString;
            }
            var partner_ids = [];
            var partner_text = [];
            var span_res = document.getElementById("partner_res")
            var partner_list = $(".partners").select2('data')
            for (var i = 0; i < partner_list.length; i++) {
                partner_ids.push(parseInt(partner_list[i].id));
                partner_text.push(partner_list[i].text);
                span_res.value = partner_text;
                span_res.innerHTML=span_res.value;
            }
            if (partner_list.length == 0){
            span_res.value = ""
            span_res.innerHTML="";
            }
            filter_data_selected.partner_ids = partner_ids

            var partner_category_ids = [];
            var partner_category_text = [];
            var span_res = document.getElementById("category_res")
            var category_list = $(".category").select2('data')

            for (var i = 0; i < category_list.length; i++) {
            if(category_list[i].element[0].selected === true)
            {partner_category_ids.push(parseInt(category_list[i].id))
            if(partner_category_text.includes(category_list[i].text) === false)
            {partner_category_text.push(category_list[i].text)
            }
            span_res.value = partner_category_text
            span_res.innerHTML=span_res.value;
            }
            }
            if (category_list.length == 0){
            span_res.value = ""
            span_res.innerHTML="";
            }
            filter_data_selected.partner_category_ids = partner_category_ids


            if ($(".target_move").length) {

            var post_res = document.getElementById("post_res")
            filter_data_selected.target_move = $(".target_move")[1].value

            post_res.value = $(".target_move")[1].value
                    post_res.innerHTML=post_res.value;
              if ($(".target_move")[1].value == "") {
              post_res.innerHTML="posted";

              }
            }

            if ($(".result_selection").length) {
            var account_res = document.getElementById("account_res")
            filter_data_selected.result_selection = $(".result_selection")[1].value
            account_res.value = $(".result_selection")[1].value
                   account_res.innerHTML=account_res.value;
              if ($(".result_selection")[1].value == "") {
              account_res.innerHTML="customer";

              }
            }


            rpc.query({
                model: 'account.partner.ageing.soa',
                method: 'write',
                args: [
                    self.wizard_id, filter_data_selected
                ],
            }).then(function(res) {
            self.initial_render = false;
                self.load_data(self.initial_render);
            });
        },

    });
    core.action_registry.add("p_a_soa", PartnerAgeingSOA);
    return PartnerAgeingSOA;
});