#using Ipopt: optimize!
using Base: Int64
using XLSX: length
#=using Pkg
Pkg.add("Tables")
Pkg.add("DataFrames")
Pkg.add("XLSX")
Pkg.add("ExcelReaders")
Pkg.add("JuMP")
Pkg.add("Ipopt")
=#
using XLSX, ExcelReaders, DataFrames, Tables, JuMP, Ipopt;
IOSource = XLSX.readdata("IO.xlsx", "io-table-5!A1:DV130");

#filepath cross system compatability code
if Sys.KERNEL === :NT || kern === :Windows
    pathmark = "\\"
else
    pathmark = "/"
end

#indexing vectors for initial data import groups
intermediaryTotalsCol = findall(x -> occursin("T4", x), string.(IOSource[3,:]));
intermediaryTotalsRow = findall(x -> occursin("T1", x), string.(IOSource[:,1]));
finalTotalsCol = findall(x -> occursin("T6", x), string.(IOSource[3,:]));
finalTotalsRow = findall(x -> occursin("Australian Production", x), string.(IOSource[:,2]));
finalDemandCol = findall(x -> occursin('Q', x), string.(IOSource[3,:]));
factorRow = findall(x -> occursin('P', x), string.(IOSource[:,1]));
IOSourceCol = vcat(intermediaryTotalsCol, finalDemandCol, finalTotalsCol);
IOSourceRow = vcat(intermediaryTotalsRow, factorRow, finalTotalsRow);
#initialising IO
IO = zeros(length(IOSourceRow), length(IOSourceCol));
#import numerical data into IO
IO[1:length(IOSourceRow), 1:length(IOSourceCol)] = IOSource[IOSourceRow, IOSourceCol];
#creating vectors of titles for IO
IOCodeRow = IOSource[IOSourceRow, 1];
IOCodeCol = IOSource[3, IOSourceCol];
IONameRow = IOSource[IOSourceRow, 2];
IONameCol = IOSource[2, IOSourceCol];

#code to sum public and private entities into one collumn
investment = findall(x -> occursin("Capital Formation", x), IONameCol);
IO[:, investment[1]]=sum(eachcol(IO[:, investment[1:2]]));
IO = IO[:,Not(investment[2])];
#alter title vectors accordingly (include Q in total investment collumn in IOcode)
IOCodeCol[investment[1]] = "Q3+Q4";
IOCodeCol = IOCodeCol[Not(investment[2])];
IONameCol[investment[1]] = "Private and Public Gross Fixed Capital Formation";
IONameCol = IONameCol[Not(investment[2])];
#creating a dictionary for the index of each collumn and row in IO by IOCode
IOColDict = Dict(IOCodeCol .=> [1:1:8;]);
IORowDict = Dict(IOCodeRow .=> [1:1:8;]);
IOCapForm = findall(x -> occursin("Capital Formation", x), IONameCol);
IOChangeInv = findall(x -> occursin("Changes in Inventories", x), IONameCol);

#importing relevant ASNA data for table 5
ASNAHouseCap = ExcelReaders.readxl("ASNAData"*pathmark*"5204039_Household_Capital_Account.xls", "Data1!A1:T71");
ASNANonFinCap = ExcelReaders.readxl("ASNAData"*pathmark*"5204018_NonFin_Corp_Capital_Account.xls", "Data1!A1:T71");
ASNAFinCap = ExcelReaders.readxl("ASNAData"*pathmark*"5204026_Fin_Corp_Capital_Account.xls", "Data1!A1:S71");
ASNAGovCap = ExcelReaders.readxl("ASNAData"*pathmark*"5204032_GenGov_Capital_Account.xls", "Data1!A1:AV71");
ASNAYearRow = findall(x -> occursin("2019", x), string.(ASNAHouseCap[:,1]));

#table 5
#creating table 5a - allocation of investment expenditure (broken into subsections for dict referencing purposes)
#subsection a is fixed capital expenditure
table5aNameCol = ["Households", "Non-Financial Corporations", "Financial Corporations", "General Government", "Total"];
table5aNameRow = ["Domestic Commodities", "Imported Commodities, complementary", "Imported Commodities, competing", 
"Taxes less subsidies on products", "Other taxes less subsidies on investment", "Total indirect taxes", 
"Total fixed capital expenditure"];
table5aRowDict = Dict(table5aNameRow .=> [1:1:length(table5aNameRow);]);
table5aColDict = Dict(table5aNameCol .=> [1:1:length(table5aNameCol);]);
table5a = zeros(length(table5aNameRow), length(table5aNameCol));

#filling in totals collumn from corresponding IO data
table5a[table5aRowDict["Domestic Commodities"], table5aColDict["Total"]] = sum(IO[IORowDict["T1"],IOCapForm]);
table5a[table5aRowDict["Imported Commodities, complementary"], table5aColDict["Total"]] = sum(IO[IORowDict["P5"],IOCapForm]);
table5a[table5aRowDict["Imported Commodities, competing"], table5aColDict["Total"]] = sum(IO[IORowDict["P6"],IOCapForm]);
table5a[table5aRowDict["Taxes less subsidies on products"], table5aColDict["Total"]] = sum(IO[IORowDict["P3"],IOCapForm]);
table5a[table5aRowDict["Other taxes less subsidies on investment"], table5aColDict["Total"]] = sum(IO[IORowDict["P4"],IOCapForm]);
table5aTaxes = findall(x -> occursin("taxes", lowercase(x)), table5aNameRow);
table5aTaxes = table5aTaxes[Not(3)];
table5a[table5aRowDict["Total indirect taxes"], table5aColDict["Total"]] = sum(table5a[table5aTaxes,table5aColDict["Total"]]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Total"]] = sum(table5a[Not(table5aRowDict["Total indirect taxes"]),table5aColDict["Total"]]);

#creating index variables for the measurements that we want
ASNAHouseCapTotCapForm = findall(x -> occursin("Gross fixed capital formation ;", x), string.(ASNAHouseCap[1,:]));
ASNANonFinCapTotCapForm = findall(x -> occursin("Gross fixed capital formation ;", x), string.(ASNANonFinCap[1,:]));
ASNAFinCapTotCapForm = findall(x -> occursin("Gross fixed capital formation ;", x), string.(ASNAFinCap[1,:]));
ASNAGenGovCapTotCapForm = findall(x -> occursin("General government ;  Gross fixed capital formation ;", x), string.(ASNAGovCap[1,:]));

#filling in totals row from ASNA Data
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Households"]] = first(ASNAHouseCap[ASNAYearRow, ASNAHouseCapTotCapForm]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Non-Financial Corporations"]] = first(ASNANonFinCap[ASNAYearRow, ASNANonFinCapTotCapForm]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Financial Corporations"]] = first(ASNAFinCap[ASNAYearRow, ASNAFinCapTotCapForm]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["General Government"]] = first(ASNAGovCap[ASNAYearRow, ASNAGenGovCapTotCapForm]);

#filling in non-total values
for ring in [1:1:length(table5aColDict)-1;];
    table5a[table5aRowDict["Domestic Commodities"],ring] = table5a[table5aRowDict["Total fixed capital expenditure"],ring]*IO[IORowDict["T1"],IOCapForm[1]] / IO[IORowDict[missing],IOCapForm[1]];
    table5a[table5aRowDict["Imported Commodities, complementary"],ring] = table5a[table5aRowDict["Total fixed capital expenditure"],ring]*IO[IORowDict["P5"],IOCapForm[1]] / IO[IORowDict[missing],IOCapForm[1]];
    table5a[table5aRowDict["Imported Commodities, competing"],ring] = table5a[table5aRowDict["Total fixed capital expenditure"],ring]*IO[IORowDict["P6"],IOCapForm[1]] / IO[IORowDict[missing],IOCapForm[1]];
    table5a[table5aRowDict["Taxes less subsidies on products"],ring] = table5a[table5aRowDict["Total fixed capital expenditure"],ring]*IO[IORowDict["P3"],IOCapForm[1]] / IO[IORowDict[missing],IOCapForm[1]];
    table5a[table5aRowDict["Other taxes less subsidies on investment"],ring] = table5a[table5aRowDict["Total fixed capital expenditure"],ring]*IO[IORowDict["P4"],IOCapForm[1]] / IO[IORowDict[missing],IOCapForm[1]];
    table5a[table5aRowDict["Total indirect taxes"], ring] = sum(table5a[table5aTaxes, ring]);
end

#creating table 5b - allocation of investment expenditure (broken into subsections for dict referencing purposes)
#subsection b is fixed capital expenditure
table5bNameCol = ["Households", "Non-Financial Corporations", "Financial Corporations", "General Government", "Total"];
table5bNameRow = ["Domestic Commodities", "Imported Commodities, complementary", "Imported Commodities, competing", 
"Taxes less subsidies on products", "Total change in inventories"];
table5bRowDict = Dict(table5bNameRow .=> [1:1:length(table5bNameRow);]);
table5bColDict = Dict(table5bNameCol .=> [1:1:length(table5bNameCol);]);
table5b = zeros(length(table5bNameRow), length(table5bNameCol));

#filling in totals collumn from corresponding IO data
table5b[table5bRowDict["Domestic Commodities"], table5bColDict["Total"]] = sum(IO[IORowDict["T1"],IOChangeInv]);
table5b[table5bRowDict["Imported Commodities, complementary"], table5bColDict["Total"]] = sum(IO[IORowDict["P5"],IOChangeInv]);
table5b[table5bRowDict["Imported Commodities, competing"], table5bColDict["Total"]] = sum(IO[IORowDict["P6"],IOChangeInv]);
table5b[table5bRowDict["Taxes less subsidies on products"], table5bColDict["Total"]] = sum(IO[IORowDict["P3"],IOChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Total"]] = sum(table5b[:,table5bColDict["Total"]]);

#creating index variables for the measurements that we want
ASNAHouseCapChangeInv = findall(x -> occursin("Changes in inventories ;", x), string.(ASNAHouseCap[1,:]));
ASNANonFinCapChangeInv = findall(x -> occursin("Changes in inventories ;", x), string.(ASNANonFinCap[1,:]));
ASNAFinCapChangeInv = findall(x -> occursin("Changes in inventories ;", x), string.(ASNAFinCap[1,:]));
ASNAGenGovCapChangeInv = findall(x -> occursin("General government ;  Changes in inventories ;", x), string.(ASNAGovCap[1,:]));

#filling in totals row from ASNA Data
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Households"]] = first(ASNAHouseCap[ASNAYearRow, ASNAHouseCapChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Non-Financial Corporations"]] = first(ASNANonFinCap[ASNAYearRow, ASNANonFinCapChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Financial Corporations"]] = first(ASNAFinCap[ASNAYearRow, ASNAFinCapChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["General Government"]] = first(ASNAGovCap[ASNAYearRow, ASNAGenGovCapChangeInv]);

#calculate non-total values with lagrangian optimisation
table5bScalingFact = abs(minimum(table5b)) * 2;
mod5b = Model(Ipopt.Optimizer);
@variable(mod5b, x[1:(length(table5bNameRow)-1), 1:(length(table5bNameCol)-1)]);
@NLobjective(mod5b, Min, sum((x[i,j] - table5bScalingFact) ^ 2 for i in 1:(length(table5bNameRow)-1), j in 1:(length(table5bNameCol)-1)));
for i in 1:(length(table5bNameRow)-1);
    @constraint(mod5b, sum(x[:,i]) == table5b[table5bRowDict["Total change in inventories"],i]+table5bScalingFact);
end;
for i in 1:(length(table5bNameCol)-1);
    @constraint(mod5b, sum(x[i,:]) == table5b[i,table5bColDict["Total"]]+table5bScalingFact);
end;
optimize!(mod5b);

#plug back into table 5b
table5b[1:(length(table5bNameRow)-1),1:(length(table5bNameCol)-1)]=value.(x).-table5bScalingFact/4;


#creating table 5c - allocation of investment expenditure (broken into subsections for dict referencing purposes)
#subsection c is totals
table5cNameCol = ["Households", "Non-Financial Corporations", "Financial Corporations", "General Government", "Total"];
table5cNameRow = ["Domestic Commodities", "Imported Commodities", "Taxes less subsidies on products", "Other taxes less subsidies on investment", "Total investment expenditure"];
table5cRowDict = Dict(table5cNameRow .=> [1:1:length(table5cNameRow);]);
table5cColDict = Dict(table5cNameCol .=> [1:1:length(table5cNameCol);]);
table5c = zeros(length(table5cNameRow), length(table5cNameCol));

#do totals calcuations to get all values in 5c
table5c[table5cRowDict["Domestic Commodities"],:] = (table5a[table5aRowDict["Domestic Commodities"],:] +
table5b[table5bRowDict["Domestic Commodities"],:]);
table5c[table5cRowDict["Imported Commodities"],:] = sum(eachrow(table5a[[table5aRowDict["Imported Commodities, competing"],table5aRowDict["Imported Commodities, complementary"]],:] +
table5b[[table5bRowDict["Imported Commodities, competing"],table5bRowDict["Imported Commodities, complementary"]],:]));
table5c[table5cRowDict["Taxes less subsidies on products"],:] = table5a[table5aRowDict["Taxes less subsidies on products"],:];
table5c[table5cRowDict["Other taxes less subsidies on investment"],:] = (table5a[table5aRowDict["Other taxes less subsidies on investment"],:] +
table5b[table5bRowDict["Taxes less subsidies on products"],:]);
table5c[table5cRowDict["Total investment expenditure"],:] = (table5a[table5aRowDict["Total fixed capital expenditure"],:] +
table5b[table5bRowDict["Total change in inventories"],:]);

#table 6
#importing relevant ASNA data
ASNAHouseInc = ExcelReaders.readxl("ASNAData"*pathmark*"5204036_Household_Income_Account.xls", "Data1!A1:AN71");
ASNANonFinInc = ExcelReaders.readxl("ASNAData"*pathmark*"5204017_NonFin_Corp_Income_Account.xls", "Data1!A1:AE71");
ASNAFinInc = ExcelReaders.readxl("ASNAData"*pathmark*"5204025_Fin_Corp_Income_Account.xls", "Data1!A1:AD71");
ASNAGovInc = ExcelReaders.readxl("ASNAData"*pathmark*"5204030_GenGov_Income_Account.xls", "Data1!A1:DA71");
ASNAExtInc = ExcelReaders.readxl("ASNAData"*pathmark*"5204043_External_Accounts.xls", "Data1!A1:AI71");
#initialising table
tableName = ["Households", "Non-Financial Corporations", "Financial Corporations", "General Government", "External", "Total"];
tableDict = Dict(tableName .=> [1:1:length(tableName);]);
table6 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
table6[tableDict["Total"],tableDict["Households"]] = (first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Interest", x), string.(ASNAHouseInc[1,:]))])
+first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Imputed interest", x), string.(ASNAHouseInc[1,:]))]));
table6[tableDict["Total"],tableDict["Non-Financial Corporations"]] = (first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("receivable - Interest", x), string.(ASNANonFinInc[1,:]))])
+first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("Property income attributed to insurance policyholders", x), string.(ASNANonFinInc[1,:]))]));
table6[tableDict["Total"],tableDict["Financial Corporations"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("receivable - Interest", x), string.(ASNAFinInc[1,:]))]);
table6[tableDict["Total"],tableDict["General Government"]] = first(ASNAGovInc[ASNAYearRow,findall(x -> occursin("General government ;  Property income receivable - Interest ;", x), string.(ASNAGovInc[1,:]))]);
table6[tableDict["Total"],tableDict["External"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("receivable - Interest", x), string.(ASNAExtInc[1,:]))]);

table6[tableDict["Households"],tableDict["Total"]] = sum(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("Property income payable - Interest", x), string.(ASNAHouseInc[1,:]))]);
table6[tableDict["Non-Financial Corporations"],tableDict["Total"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("Property income payable - Interest", x), string.(ASNANonFinInc[1,:]))]);
table6[tableDict["Financial Corporations"],tableDict["Total"]] = (first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("Property income payable - Interest", x), string.(ASNAFinInc[1,:]))])
+first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("payable - Property income attributed to insurance policy holders", x), string.(ASNAFinInc[1,:]))]));
table6[tableDict["General Government"],tableDict["Total"]] = first(ASNAGovInc[ASNAYearRow,findall(x -> occursin("General government ;  Property income payable - Total interest ;", x), string.(ASNAGovInc[1,:]))]);
table6[tableDict["External"],tableDict["Total"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("Property income payable - Interest", x), string.(ASNAExtInc[1,:]))]);

if 0.98*sum(table6[:,length(tableName)])<sum(table6[length(tableName),:])<1.02*sum(table6[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 6")
end

#assuming gov only receives interest from fin corps
table6[tableDict["Financial Corporations"], tableDict["General Government"]] = table6[tableDict["Total"], tableDict["General Government"]];
#from other ASNA data
table6[tableDict["Financial Corporations"], tableDict["Non-Financial Corporations"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("attributed", x), string.(ASNANonFinInc[1,:]))]);
table6[tableDict["Financial Corporations"], tableDict["Households"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("attributed", x), string.(ASNAFinInc[1,:]))])-first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("attributed", x), string.(ASNANonFinInc[1,:]))]);
table6[tableDict["General Government"], tableDict["Households"]] = first(ASNAGovInc[ASNAYearRow,findall(x -> occursin("General government ;  Property income payable - Interest - On unfunded superannuation liabilities ;", x), string.(ASNAGovInc[1,:]))]);

#solve for missing values with scaling
table6Step3 = zeros(length(tableName),length(tableName));
table6Step3Row = [tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"],tableDict["General Government"]];
table6Step3Col = [tableDict["External"]];
for i in table6Step3Col;
    for ring in table6Step3Row;
        table6Step3[ring,i] = (table6[tableDict["Total"],i]-sum(table6[1:(length(tableName)-1),i]))*(
            table6[ring,tableDict["Total"]]-sum(table6[ring,1:(length(tableName)-1)]))/sum(table6[
            table6Step3Row,tableDict["Total"]]-sum(eachcol(table6[table6Step3Row,1:(length(tableName)-1)])));
    end
end
table6 = table6+table6Step3;

table6Step4 = zeros(length(tableName),length(tableName));
table6Step4Row = [1:1:(length(tableName)-1);];
table6Step4Col = [tableDict["Households"],tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"]];
for i in table6Step4Col;
    for ring in table6Step4Row;
        table6Step4[ring,i] = (table6[tableDict["Total"],i]-sum(table6[1:(length(tableName)-1),i]))*(
            table6[ring,tableDict["Total"]]-sum(table6[ring,1:(length(tableName)-1)]))/sum(table6[
            table6Step4Row,tableDict["Total"]]-sum(eachcol(table6[table6Step4Row,1:(length(tableName)-1)])));
    end
end
table6 = table6+table6Step4;

#table 7
#initialising table
table7 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
table7[tableDict["Total"],tableDict["Households"]] = first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Dividends", x), string.(ASNAHouseInc[1,:]))]);
table7[tableDict["Total"],tableDict["Non-Financial Corporations"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("receivable - Dividends", x), string.(ASNANonFinInc[1,:]))]);
table7[tableDict["Total"],tableDict["Financial Corporations"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("receivable - Dividends", x), string.(ASNAFinInc[1,:]))]);
table7[tableDict["Total"],tableDict["General Government"]] = first(ASNAGovInc[ASNAYearRow,findall(x -> occursin("General government ;  Property income receivable - Dividends ;", x), string.(ASNAGovInc[1,:]))]);
table7[tableDict["Total"],tableDict["External"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("receivable - Dividends", x), string.(ASNAExtInc[1,:]))]);

table7[tableDict["Households"],tableDict["Total"]] = 0.0
table7[tableDict["Non-Financial Corporations"],tableDict["Total"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("payable - Dividends", x), string.(ASNANonFinInc[1,:]))]);
table7[tableDict["Financial Corporations"],tableDict["Total"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("payable - Dividends", x), string.(ASNAFinInc[1,:]))]);
table7[tableDict["General Government"],tableDict["Total"]] = 0.0;
table7[tableDict["External"],tableDict["Total"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("payable - Dividends", x), string.(ASNAExtInc[1,:]))]);

if 0.98*sum(table7[:,length(tableName)])<sum(table7[length(tableName),:])<1.02*sum(table7[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 7")
end

#assuming gov only receives dividends from non-fin corps
table7[tableDict["Non-Financial Corporations"], tableDict["General Government"]] = table7[tableDict["Total"], tableDict["General Government"]];

#solve for missing values with scaling
table7Step3 = zeros(length(tableName),length(tableName));
table7Step3Row = [tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"]];
table7Step3Col = [tableDict["External"]];
for i in table7Step3Col;
    for ring in table7Step3Row;
        table7Step3[ring,i] = (table7[tableDict["Total"],i]-sum(table7[1:(length(tableName)-1),i]))*(
            table7[ring,tableDict["Total"]]-sum(table7[ring,1:(length(tableName)-1)]))/sum(table7[
            table7Step3Row,tableDict["Total"]]-sum(eachcol(table7[table7Step3Row,1:(length(tableName)-1)])));
    end
end
table7 = table7+table7Step3;

table7Step4 = zeros(length(tableName),length(tableName));
table7Step4Row = [tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"], tableDict["External"]];
table7Step4Col = [tableDict["Households"],tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"]];
for i in table7Step4Col;
    for ring in table7Step4Row;
        table7Step4[ring,i] = (table7[tableDict["Total"],i]-sum(table7[1:(length(tableName)-1),i]))*(
            table7[ring,tableDict["Total"]]-sum(table7[ring,1:(length(tableName)-1)]))/sum(table7[
            table7Step4Row,tableDict["Total"]]-sum(eachcol(table7[table7Step4Row,1:(length(tableName)-1)])));
    end
end
table7 = table7+table7Step4;

#table 8
#initialising table
table8 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
table8[tableDict["Total"],tableDict["Households"]] = first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Reinvested", x), string.(ASNAHouseInc[1,:]))]);
table8[tableDict["Total"],tableDict["Non-Financial Corporations"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("receivable - Reinvested", x), string.(ASNANonFinInc[1,:]))]);
table8[tableDict["Total"],tableDict["Financial Corporations"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("receivable - Reinvested", x), string.(ASNAFinInc[1,:]))]);
table8[tableDict["Total"],tableDict["General Government"]] = 0.0;
table8[tableDict["Total"],tableDict["External"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("receivable - Reinvested", x), string.(ASNAExtInc[1,:]))]);

table8[tableDict["Households"],tableDict["Total"]] = 0.0
table8[tableDict["Non-Financial Corporations"],tableDict["Total"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("payable - Reinvested", x), string.(ASNANonFinInc[1,:]))]);
table8[tableDict["Financial Corporations"],tableDict["Total"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("payable - Reinvested", x), string.(ASNAFinInc[1,:]))]);
table8[tableDict["General Government"],tableDict["Total"]] = 0.0;
table8[tableDict["External"],tableDict["Total"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("payable - Reinvested", x), string.(ASNAExtInc[1,:]))]);

if 0.98*sum(table8[:,length(tableName)])<sum(table8[length(tableName),:])<1.02*sum(table8[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 8")
end
#assuming that values can be negative
#solve missing values with Ipopt
mod8 = Model(Ipopt.Optimizer);
@variable(mod8, x[1:(length(tableName)-1), 1:(length(tableName)-1)]);
@NLobjective(mod8, Min, sum((x[i,j]) ^ 2 for i in 1:(length(tableName)-1), j in 1:(length(tableName)-1)));
for i in 1:(length(tableName)-1);
    @constraint(mod8, sum(x[:,i]) == table8[tableDict["Total"],i]-sum(table8[1:(length(tableName)-1),i]));
    @constraint(mod8, sum(x[i,:]) == table8[i,tableDict["Total"]]-sum(table8[i,1:(length(tableName)-1)]));
    @constraint(mod8, x[tableDict["General Government"],i] == 0);
    @constraint(mod8, x[i,tableDict["General Government"]] == 0);
end;
@constraint(mod8, x[tableDict["External"],tableDict["External"]] == 0);
@constraint(mod8, x[tableDict["Households"],tableDict["Households"]] == 0);
optimize!(mod8);
table8[1:(length(tableName)-1), 1:(length(tableName)-1)] = table8[1:(length(tableName)-1), 1:(length(tableName)-1)] + value.(x);


#=spread the external receivable totals between fin and non-fin Corporations
table8Step3 = zeros(length(tableName),length(tableName));
table8Step3Row = [tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"]];
table8Step3Col = [tableDict["External"]];
for i in table8Step3Col;
    for ring in table8Step3Row;
        table8Step3[ring,i] = (table8[tableDict["Total"],i]-sum(table8[1:(length(tableName)-1),i]))*(
            table8[ring,tableDict["Total"]]-sum(table8[ring,1:(length(tableName)-1)]))/sum(table8[
            table8Step3Row,tableDict["Total"]]-sum(eachcol(table8[table8Step3Row,1:(length(tableName)-1)])));
    end
end
table8 = table8+table8Step3;
=#

#=table 9
#initialising table
table9 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
table9[tableDict["Total"],tableDict["Households"]] = first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Rent on natural", x), string.(ASNAHouseInc[1,:]))]);
table9[tableDict["Total"],tableDict["Non-Financial Corporations"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("receivable - Rent on natural", x), string.(ASNANonFinInc[1,:]))]);
table9[tableDict["Total"],tableDict["Financial Corporations"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("receivable - Rent on natural", x), string.(ASNAFinInc[1,:]))]);
table9[tableDict["Total"],tableDict["General Government"]] = first(ASNAGovInc[ASNAYearRow,findall(x -> occursin("receivable - Rent on natural", x), string.(ASNAGovInc[1,:]))]);
table9[tableDict["Total"],tableDict["External"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("receivable - Reinvested", x), string.(ASNAExtInc[1,:]))]);

table9[tableDict["Households"],tableDict["Total"]] =  first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("payable - Rent on natural", x), string.(ASNAHouseInc[1,:]))]);
table9[tableDict["Non-Financial Corporations"],tableDict["Total"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("payable - Rent on natural", x), string.(ASNANonFinInc[1,:]))]);
table9[tableDict["Financial Corporations"],tableDict["Total"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("payable - Rent on natural", x), string.(ASNAFinInc[1,:]))]);
table9[tableDict["General Government"],tableDict["Total"]] =  first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Rent on natural", x), string.(ASNAHouseInc[1,:]))]);
table9[tableDict["External"],tableDict["Total"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("payable - Reinvested", x), string.(ASNAExtInc[1,:]))]);

if 0.98*sum(table9[:,length(tableName)])<sum(table9[length(tableName),:])<1.02*sum(table9[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 9")
end

#spread the external receivable totals between fin and non-fin Corporations
table9Step3 = zeros(length(tableName),length(tableName));
table9Step3Row = [tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"],tableDict["External"]];
table9Step3Col = [tableDict["Households"],tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"],tableDict["External"]];
for i in table9Step3Col;
    for ring in table9Step3Row;
        table9Step3[ring,i] = (table9[tableDict["Total"],i]-sum(table9[1:(length(tableName)-1),i]))*(
            table9[ring,tableDict["Total"]]-sum(table9[ring,1:(length(tableName)-1)]))/sum(table9[
            table9Step3Row,tableDict["Total"]]-sum(eachcol(table9[table9Step3Row,1:(length(tableName)-1)])));
    end
end
table9 = table9+table9Step3;
=#


#=convert dataframe to dictionary
function increment!( d::Dict{S, T}, k::S, i::T) where {T<:Real, S<:Any}
    if haskey(d, k)
        d[k] += i
    else
        d[k] = i
    end
end
increment!(d::Dict{S, T}, k::S ) where {T<:Real, S<:Any} = increment!( d, k, one(T))

function df2dict( df::DataFrame, key_col::Symbol, val_col::Symbol=:null)
    keytype = typeof(df[1,key_col])
    if val_col == :null
        valtype = Int
    else
        valtype = typeof(df[1,val_col])
    end
    D = Dict{keytype, valtype}()
    for i=1:size(df,1)
        if !ismissing(df[i,key_col])
            if val_col == :null
                increment!( D, df[i,key_col] )
            elseif valtype <: Real
                increment!( D, df[i,key_col], df[i,val_col] )
            else
                if haskey(D, df[i,key_col])
                    @warn("non-unique entry: $(df[i,key_col])")
                else
                    D[df[i,key_col]] = df[i,val_col]
                end
            end
        end
    end
    return D
end
df[!, "IOCode"]=IOcode
insertcols!(df, 2, :name => vector)
D = df2dict(df, :IOCode, :x3)
=#