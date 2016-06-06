.. _cht-multilvl-cat:

Chart - Multilevel Categories
=============================

A category chart can have more than one level of categories, where each
individual category is nested in a higher-level category. There can be
multiple levels.

https://www.youtube.com/watch?v=KbyQpzA7fLo


XML specimens
-------------

.. highlight:: xml

::

  <c:cat>
     <c:multiLvlStrRef>
       <c:f>Sheet1!$C$1:$J$3</c:f>
       <c:multiLvlStrCache>
         <c:ptCount val="8"/>
         <c:lvl>
           <c:pt idx="0">
             <c:v>county one</c:v>
           </c:pt>
           <c:pt idx="1">
             <c:v>county two</c:v>
           </c:pt>
           <c:pt idx="2">
             <c:v>county one</c:v>
           </c:pt>
           <c:pt idx="3">
             <c:v>county two</c:v>
           </c:pt>
           <c:pt idx="4">
             <c:v>county one</c:v>
           </c:pt>
           <c:pt idx="5">
             <c:v>county two</c:v>
           </c:pt>
           <c:pt idx="6">
             <c:v>country one</c:v>
           </c:pt>
           <c:pt idx="7">
             <c:v>county two</c:v>
           </c:pt>
         </c:lvl>
         <c:lvl>
           <c:pt idx="0">
             <c:v>city one</c:v>
           </c:pt>
           <c:pt idx="2">
             <c:v>city two </c:v>
           </c:pt>
           <c:pt idx="4">
             <c:v>city one</c:v>
           </c:pt>
           <c:pt idx="6">
             <c:v>city two </c:v>
           </c:pt>
         </c:lvl>
         <c:lvl>
           <c:pt idx="0">
             <c:v>UK</c:v>
           </c:pt>
           <c:pt idx="4">
             <c:v>US</c:v>
           </c:pt>
         </c:lvl>
       </c:multiLvlStrCache>
     </c:multiLvlStrRef>
   </c:cat>


  <c:catAx>
    ...
    <c:noMultiLvlLbl val="0"/>
  </c:catAx>
