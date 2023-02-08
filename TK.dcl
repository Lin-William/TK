TK:dialog {
  label="TKmaker";
  : column {
  : edit_box { label="abbr"; edit_width =10; alignment =left; value ="W2022212";key="kabbr";}
  : edit_box { label="总图"; edit_width =10; value="1";key="kzt";}
  : edit_box { label="平面"; edit_width =10;key="kpm"; }
  : edit_box { label="纵断"; edit_width =10; key="kzd";}
  : edit_box { label="材料"; edit_width =10; value="1";key="kcl";}
  : edit_box { label="横断"; edit_width =10; key="khd";}
  : edit_box { label="节点"; edit_width =10; value="1";key="kjd";}
  children_fixed_width=true;
  }
  spacer_1;
  ok_cancel;
}