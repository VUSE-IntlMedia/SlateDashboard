import streamlit as st
import pandas as pd
import numpy as np
import re
import pycountry_convert as pc
import plotly.express as px
import plotly.colors


@st.cache_data
def load_data_from_excel(file):
    us_states = ['Alabama', 'Alaska', 'Arizona', 'Arkansas', 'California', 'Colorado', 'Connecticut', 'Delaware',
                 'Florida', 'Georgia', 'Hawaii', 'Idaho', 'Illinois', 'Indiana', 'Iowa', 'Kansas', 'Kentucky',
                 'Louisiana', 'Maine', 'Maryland', 'Massachusetts', 'Michigan', 'Minnesota', 'Mississippi', 'Missouri',
                 'Montana', 'Nebraska', 'Nevada', 'New Hampshire', 'New Jersey', 'New Mexico', 'New York',
                 'North Carolina', 'North Dakota', 'Ohio', 'Oklahoma', 'Oregon', 'Pennsylvania', 'Rhode Island',
                 'South Carolina', 'South Dakota', 'Tennessee', 'Texas', 'Utah', 'Vermont', 'Virginia', 'Washington',
                 'West Virginia', 'Wisconsin', 'Wyoming', 'Puerto Rico', 'United States']
    us_insts = ['American University', 'Amherst College', 'Appalachian State University', 'Auburn University',
                'Baylor University', 'Boston College', 'Boston University', 'Brandeis University', 'Brown University',
                'Bucknell University', 'Carnegie Mellon University', 'Case Western Reserve University',
                'Clark University', 'Clarkson University', 'Clemson University', 'Colby College',
                'College of William and Mary', 'Columbia University', 'Cornell University', 'Dartmouth College',
                'DePaul University', 'Drexel University', 'Duke University', 'Duquesne University', 'Elon University',
                'Emory University', 'Fisk University', 'George Mason University', 'Georgetown University',
                'Hanover College', 'Harvard University', 'Howard University', 'Johns Hopkins University',
                'Kennesaw State University', 'Lafayette College', 'Lehigh University', 'Lipscomb University',
                'Luther College', 'Mercer University', 'Miami University', 'Millikin University',
                'Milwaukee School of Engineering', 'Mount Holyoke College', 'Murray State University',
                'Norfolk State University', 'Prairie View A & M University', 'Princeton University',
                'Purdue University', 'Rensselaer Polytechnic Institute', 'Rice University', 'Roger Williams University',
                'Rutgers University', 'SUNY Binghamton', 'SUNY Buffalo', 'SUNY College', 'SUNY Empire',
                'SUNY Polytechnic Institute', 'SUNY Stony Brook', 'SUNY University', 'Saint Louis University',
                'Stanford University', 'Stetson University', 'Stevens Institute of Technology', 'Stony Brook University',
                'Stratford University', 'Syracuse University', 'Temple University', 'Trine University',
                'Trinity University', 'Tufts University', 'Tulane University', 'Union College', 'University at Buffalo',
                'University of Chicago', 'University of Cincinnati', 'University of Dayton', 'University of Houston',
                'University of Louisville', 'University of Memphis', 'University of Miami', 'University of Notre Dame',
                'University of Phoenix', 'University of Pittsburgh', 'University of Rochester', 'Vanderbilt University',
                'Villanova University', 'Wake Forest University', 'Wesleyan University', 'Westminster College',
                'Wheaton College', 'Wofford College', 'Worcester Polytechnic Institute', 'Wright State University',
                'Yale University', 'Yeshiva University']
    specials = ['Northeastern University', 'Northwestern University', 'Southwestern University', 'Southeastern University']

    dataframe = pd.read_excel(file)
    dataframe["Program"] = dataframe["Program"].replace({
        "Electrical Engineering": "Electrical and Computer Engineering"
    })
    dataframe["Degree"] = dataframe["Degree"].replace({
        "ME": "MS/ME",
        "MS": "MS/ME"
    })
    dataframe["Year"] = dataframe["Application Term"].str.slice(0, 4)
    dataframe["Citizenship2"] = np.where(dataframe["Citizenship2"] == "United States", dataframe["Citizenship1"], dataframe["Citizenship2"])
    dataframe["Citizenship1"] = np.where(dataframe["Citizenship"] == "US", "United States", dataframe["Citizenship1"])
    dataframe["Continent"] = dataframe["Citizenship1"].apply(get_continent)
    dataframe["App Submitted"] = pd.to_datetime(dataframe["App Submitted"]).dt.date
    dataframe["Submission Month"] = pd.to_datetime(dataframe["App Submitted"]).dt.strftime("%Y %b")
    dataframe["is_usuni"] = np.where(dataframe["School 1 Institution"].str.contains(
        rf"({'|'.join(us_states + us_insts)})", regex=True, case=False) | (
            (dataframe["School 1 Institution"].str.contains(rf"({'|'.join(specials)})", regex=True, case=False)) & (
                dataframe["Active Country"] == "United States")), "yes", "no")

    if "Decision Most Recent Confirmed Name" in dataframe.columns:
        dataframe.rename(columns={"Decision Most Recent Confirmed Name": "Decision"}, inplace=True)

    return dataframe[~dataframe["Degree"].isin([np.nan, "", "Non-Degree"])]


@st.cache_data
def get_year_from_data(dataframe: pd.DataFrame):
    if "Application Term" in dataframe.columns:
        return list(sorted(dataframe["Year"].dropna().unique(), reverse=True))
    else:
        return ["'Application Term' Column Not Found"]


@st.cache_data
def get_decision_from_data(dataframe: pd.DataFrame):
    if "Decision" in dataframe.columns:
        return list(sorted(set(dataframe["Decision"].dropna().unique()).difference({"", None})))
    else:
        return ["'Decision' Column Not Found"]


@st.cache_data
def add_filters(dataframe: pd.DataFrame, year, program):
    year_list, program_list = year, program
    if isinstance(year, (str, int, float)):
        year_list = [year]
    if isinstance(program, (str, int, float)):
        program_list = [program]
    if (not dataframe.empty) and any([year, program]):
        return dataframe[(dataframe["Year"].astype("str").isin([str(y) for y in year_list])) & (
            dataframe["Program"].isin(program_list))]
    else:
        return dataframe


@st.cache_data
def create_details_submit(dataframe: pd.DataFrame, years: list, programs: list, degrees: list):
    details_submit = {}
    for y in years:
        details_submit.update({y: {}})
        details_submit[y].update(
            {"all": {"all": dataframe[
                dataframe["Year"].astype(str) == str(y)
            ]["Status"].dropna().value_counts().to_dict()}})
        for d in degrees:
            details_submit[y]["all"].update(
                {d: dataframe[
                    (dataframe["Year"].astype(str) == str(y)) & (dataframe["Degree"] == d)
                    ]["Status"].dropna().value_counts().to_dict()})
        for p in programs:
            details_submit[y].update({p: {}})
            details_submit[y][p].update(
                {"all": dataframe[
                    (dataframe["Year"].astype(str) == str(y)) & (dataframe["Program"] == p)
                    ]["Status"].dropna().value_counts().to_dict()})
            for d in degrees:
                details_submit[y][p].update(
                    {d: dataframe[
                        (dataframe["Year"].astype(str) == str(y)) & (dataframe["Program"] == p) & (dataframe["Degree"] == d)
                    ]["Status"].dropna().value_counts().to_dict()})
    return details_submit


@st.cache_data
def create_details_decide(df_decide: pd.DataFrame, years: list, programs: list, degrees: list):
    details_decide = {}
    df_decide = df_decide[df_decide["Status"] == "Decided"]
    for y in years:
        details_decide.update({y: {}})
        details_decide[y].update(
            {"all": {"all": df_decide[
                df_decide["Year"].astype(str) == str(y)
                ]["Decision"].dropna().value_counts().to_dict()}})
        for d in degrees:
            details_decide[y]["all"].update(
                {d: df_decide[
                    (df_decide["Year"].astype(str) == str(y)) & (df_decide["Degree"] == d)
                    ]["Decision"].dropna().value_counts().to_dict()})
        for p in programs:
            details_decide[y].update({p: {}})
            details_decide[y][p].update(
                {"all": df_decide[
                    (df_decide["Year"].astype(str) == str(y)) & (df_decide["Program"] == p)
                    ]["Decision"].dropna().value_counts().to_dict()})
            for d in degrees:
                details_decide[y][p].update(
                    {d: df_decide[
                        (df_decide["Year"].astype(str) == str(y)) & (df_decide["Program"] == p) & (df_decide["Degree"] == d)
                        ]["Decision"].dropna().value_counts().to_dict()})
    return details_decide


def sum_dict(dic, keys=None):
    if not keys:
        return sum(dic.values())
    elif isinstance(keys, str):
        return dic[keys]
    elif not dic:
        return 0
    else:
        return sum({k: dic[k] for k in keys if k in dic}.values())


def get_continent(country_name):
    missing_territories = {
        "Congo (Kinshasa)": "Africa",
        "The Bahamas": "North America",
        "Cote D'Ivoire": "Africa",
        "Macau S.A.R.": "Asia",
        "Glorioso Islands": "Africa",
        "Ashmore and Cartier Islands": "Oceania",
        "Hong Kong S.A.R.": "Asia",
        "Congo (Brazzaville)": "Africa",
        "The Gambia": "Africa"
    }
    if country_name in missing_territories:
        return missing_territories[country_name]
    elif country_name not in (np.nan, ""):
        try:
            country_code = pc.country_name_to_country_alpha2(country_name)
            continent_code = pc.country_alpha2_to_continent_code(country_code)
            return pc.convert_continent_code_to_continent_name(continent_code)
        except:
            return "Unknown"


def get_top(series: pd.Series, index_name: str, num: int):
    tops = series.value_counts().head(num).reset_index()
    tops.columns = [index_name, "Count"]
    tops["Percentage"] = (tops["Count"] / tops["Count"].sum() * 100).round(2)
    return tops


def delta_params(delta, reverse=False):
    if reverse:
        return ('↓', '#FF0000', '#FFE8E8') if delta > 0 else ('→', '#6E6E6E', '#F7F7F7') if delta == 0 else (
            '↑', '#0F6615', '#EBFCEC')
    else:
        return ('↓', '#FF0000', '#FFE8E8') if delta < 0 else ('→', '#6E6E6E', '#F7F7F7') if delta == 0 else (
            '↑', '#0F6615', '#EBFCEC')


def convert_rgb(rgb_text, color_type):
    if color_type == "rgba":
        return rgb_text.replace("rgb(", "rgba(").replace(")", f", 0.5)")
    elif color_type == "hex":
        rgb_values = list(map(int, re.findall(r"\d+", rgb_text)))
        return "#{:02X}{:02X}{:02X}".format(*rgb_values)


def calc_metrics(data_dict, year1, year2, program, category, keys):
    curr = sum_dict(data_dict[year1][program][category], keys)
    last = sum_dict(data_dict[year2][program][category], keys)
    delta = (curr / last - 1) * 100 if last else 0
    params = delta_params(delta)
    return curr, last, delta, params


colors = ["curl", "BrBG", "PRGn", "delta", "Armyrose", "Fall", "Tealrose", "Tropic", "Earth"]
curr_yr, swapped, program_selected, degree_selected = 0, None, [], []
st.set_page_config(
    page_title="Admission Insights",
    layout="wide",
    initial_sidebar_state="collapsed"
)

with st.sidebar:
    st.markdown(f"""
        <p style="font-size: 12px; color: #B6B6B6; margin-top: -15px; margin-bottom: 20px; text-align: center;">
            &copy; 2025 - Dongying Tao
            <br>
            All rights reserved.
            <br>
            Version Date: Aug 26, 2025
        </p>""", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])
    if uploaded_file:
        df = load_data_from_excel(uploaded_file)
        year_options = get_year_from_data(df)
        program_options = ['Biomedical Engineering', 'Chemical Engineering', 'Civil Engineering', 'Computer Science',
                           'Cyber Physical Systems', 'Electrical and Computer Engineering',
                           'Engineering in Surgery and Intervention', 'Environmental Engineering',
                           'Interdisciplinary Materials Science', 'Mechanical Engineering',
                           'Risk, Reliability, and Resilience']  # 'Engineering Non-Degree'
        degree_options = ["MS/ME", "PHD"]  # "Non-Degree"
        decision_options = get_decision_from_data(df)

        submitted_keys = ["Awaiting Decision", "Decided"]
        offered_keys = ["Admit", "Intention to Matriculate"]
        intended_keys = ["Intention to Matriculate"]

        submits = create_details_submit(df, years=year_options, programs=program_options, degrees=degree_options)
        decides = create_details_decide(df, years=year_options, programs=program_options, degrees=degree_options)

        year_all_on = st.toggle("All Years", value=True)
        if not year_all_on:
            year_options_defined = [y for i, y in enumerate(year_options) if i != len(year_options)-1]
            curr_yr = st.pills("Year:", year_options_defined, selection_mode="single", label_visibility="collapsed")
        if not curr_yr:
            curr_yr = max(year_options)
        last_yr = str(int(curr_yr) - 1)

        program_all_on = st.toggle("All Programs", value=True)
        if not program_all_on:
            program_selected = st.pills("Program:", program_options, selection_mode="single", label_visibility="collapsed")
        if not program_selected:
            program_selected = program_options
        df_submitted = add_filters(df[df["Status"].isin(submitted_keys)], year=curr_yr, program=program_selected)
        df_intended = add_filters(df[df["Decision"].isin(intended_keys)], year=curr_yr, program=program_selected)
    else:
        no_data_on = st.toggle("No Data Loaded", value=False, disabled=True)

if curr_yr:
    # TOP: Header & Facts
    with st.container():
        # Header
        top_header, top_color, top_reverse = st.columns([4, 1, 0.2])
        program_text = "VUSE All Programs" if isinstance(program_selected, list) else program_selected
        top_header.markdown(f"""
            <p style='font-size:23px; line-height:1; margin-bottom:10px; text-align:left;'>
            <strong>{curr_yr} Postgraduate Admissions Summary | {program_text}</strong>
            </p>
        """, unsafe_allow_html=True)
        with top_color:
            colorscale = st.selectbox(
                "Colors",
                colors,
                index=0,
                label_visibility="collapsed",
                placeholder="Choose your preferred color theme..."
            )
        with top_reverse:
            color_reversed = st.pills(
                "Swapped",
                options=[":material/swap_vert:"],
                selection_mode="single",
                label_visibility="collapsed"
            )
            if color_reversed:
                colorscale = f"{colorscale}_r"
        rgb_picked = [rgb for i, rgb in plotly.colors.get_colorscale(colorscale)]

        # Facts
        top_submit, top_offer, top_intent, top_progress = st.columns([1, 1, 1, 0.5])
        top_height = 150
        program_key = program_selected if isinstance(program_selected, str) else "all"

        sub_curr, sub_last, sub_delta, sub_pms = calc_metrics(submits, curr_yr, last_yr, program_key, "all", submitted_keys)
        off_curr, off_last, off_delta, off_pms = calc_metrics(decides, curr_yr, last_yr, program_key, "all", offered_keys)
        int_curr, int_last, int_delta, int_pms = calc_metrics(decides, curr_yr, last_yr, program_key, "all", intended_keys)
        msub_curr, msub_last, msub_delta, msub_pms = calc_metrics(submits, curr_yr, last_yr, program_key, "MS/ME", submitted_keys)
        moff_curr, moff_last, moff_delta, moff_pms = calc_metrics(decides, curr_yr, last_yr, program_key, "MS/ME", offered_keys)
        mint_curr, mint_last, mint_delta, mint_pms = calc_metrics(decides, curr_yr, last_yr, program_key, "MS/ME", intended_keys)
        psub_curr, psub_last, psub_delta, psub_pms = calc_metrics(submits, curr_yr, last_yr, program_key, "PHD", submitted_keys)
        poff_curr, poff_last, poff_delta, poff_pms = calc_metrics(decides, curr_yr, last_yr, program_key, "PHD", offered_keys)
        pint_curr, pint_last, pint_delta, pint_pms = calc_metrics(decides, curr_yr, last_yr, program_key, "PHD", intended_keys)

        with top_submit.container(border=True, height=top_height):
            submit1, submit2 = st.columns([1.5, 1])
            submit1.markdown(f"""
                <p style='font-size:14px; line-height:0; margin-top:5px; text-align:left; color:#0D0D0D;'>
                <strong>Submissions Completed</strong>
                </p>
                
                <p style='font-size:36px; line-height:1; margin-top:10px; text-align:left; color:#0D0D0D;'>
                {sub_curr:,}
                </p>
                
                <div style='display: flex; justify-content: space-between; line-height:1; margin-top:-5px;'>
                    <span style='color:{sub_pms[1]}; background-color:{sub_pms[2]}; 
                                 padding:4px 6px; border-radius:10px; font-size:13px;'>
                    {sub_pms[0]} {sub_delta:,.0f} %
                    </span>
                    <span style='font-size:13px; background-color:#FAFAFA;
                                 padding:4px 8px; border-radius:5px; color:#808080;'>
                    <i><strong>Prior: {sub_last:,}</strong></i>
                    </span>
                </div>
            """, unsafe_allow_html=True)
            with submit2:
                st.markdown(f"""
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div style="display:flex; flex-direction:column; align-items:flex-start; justify-content:space-between;">
                            <div style="font-size:11px; line-height:1; margin:0px; color:#0D0D0D;">
                                <strong>MS/ME</strong>
                            </div>
                            <div style="font-size:24px; line-height:1.2; color:#0D0D0D; ">
                                {msub_curr:,}
                            </div>
                        </div>
                        <div style="display:flex; flex-direction:column; align-items:flex-end; justify-content:space-between;">
                            <span style="font-size:11px; color:{msub_pms[1]}; background-color:{msub_pms[2]};
                                        padding:1px 4px; border-radius:10px; text-align:right; margin:-2px; ">
                                {msub_pms[0]} {msub_delta:,.0f} %
                            </span>
                            <p style="font-size:10px; line-height:1; margin-top:8px; text-align:left; color:#808080;">
                                <i><strong>Prior: {msub_last:,}</strong></i>
                            </p>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                st.markdown(f"""
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div style="display:flex; flex-direction:column; align-items:flex-start; justify-content:space-between;">
                            <div style="font-size:11px; line-height:1; margin:0px; color:#0D0D0D;">
                                <strong>PHD</strong>
                            </div>
                            <div style="font-size:24px; line-height:1.2; color:#0D0D0D; ">
                                {psub_curr:,}
                            </div>
                        </div>
                        <div style="display:flex; flex-direction:column; align-items:flex-end; justify-content:space-between;">
                            <span style="font-size:11px; color:{psub_pms[1]}; background-color:{psub_pms[2]};
                                        padding:1px 4px; border-radius:10px; margin:-2px; ">
                                {psub_pms[0]} {psub_delta:,.0f} %
                            </span>
                            <p style="font-size:10px; line-height:1; margin-top:10px; color:#808080;">
                                <i><strong>Prior: {psub_last:,}</strong></i>
                            </p>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
            st.markdown(f"""
                <hr style='margin-top: 3px; margin-bottom: 5px;'>
                <p style='font-size:10px; line-height:1.2; margin-top:0px; margin-bottom:1px;'>
                <strong>Status Labels</strong>: {'; '.join(submitted_keys)}
                </p>
            """, unsafe_allow_html=True)
        with top_offer.container(border=True, height=top_height):
            offer1, offer2 = st.columns([1.5, 1])
            offer1.markdown(f"""
                <p style='font-size:14px; line-height:0; margin-top:5px; text-align:left; color:#0D0D0D;'>
                <strong>Offers Issued</strong>
                </p>
                
                <p style='font-size:36px; line-height:1; margin-top:10px; text-align:left; color:#0D0D0D;'>
                {off_curr:,}
                </p>
                
                <div style='display: flex; justify-content: space-between; line-height:1; margin-top:-5px;'>
                    <span style='color:{off_pms[1]}; background-color:{off_pms[2]}; 
                                 padding:4px 6px; border-radius:10px; font-size:13px;'>
                     {off_pms[0]} {off_delta:,.0f} %
                    </span>
                    <span style='font-size:13px; background-color:#FAFAFA;
                                 padding:4px 8px; border-radius:5px; color:#808080;'>
                     <i><strong>Prior: {off_last:,}</strong></i>
                    </span>
                </div>
            """, unsafe_allow_html=True)
            with offer2:
                st.markdown(f"""
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div style="display:flex; flex-direction:column; align-items:flex-start; justify-content:space-between;">
                            <div style="font-size:11px; line-height:1; margin:0px; color:#0D0D0D;">
                                <strong>MS/ME</strong>
                            </div>
                            <div style="font-size:24px; line-height:1.2; color:#0D0D0D; ">
                                {moff_curr:,}
                            </div>
                        </div>
                        <div style="display:flex; flex-direction:column; align-items:flex-end; justify-content:space-between;">
                            <span style="font-size:11px; color:{moff_pms[1]}; background-color:{moff_pms[2]};
                                        padding:1px 4px; border-radius:10px; margin:-2px; ">
                                {moff_pms[0]} {moff_delta:,.0f} %
                            </span>
                            <p style="font-size:10px; line-height:1; margin-top:10px; color:#808080;">
                                <i><strong>Prior: {moff_last:,}</strong></i>
                            </p>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                st.markdown(f"""
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div style="display:flex; flex-direction:column; align-items:flex-start; justify-content:space-between;">
                            <div style="font-size:11px; line-height:1; margin:0px; color:#0D0D0D;">
                                <strong>PHD</strong>
                            </div>
                            <div style="font-size:24px; line-height:1.2; color:#0D0D0D; ">
                                {poff_curr:,}
                            </div>
                        </div>
                        <div style="display:flex; flex-direction:column; align-items:flex-end; justify-content:space-between;">
                            <span style="font-size:11px; color:{poff_pms[1]}; background-color:{poff_pms[2]};
                                        padding:1px 4px; border-radius:10px; margin:-2px; ">
                                {poff_pms[0]} {poff_delta:,.0f} %
                            </span>
                            <p style="font-size:10px; line-height:1; margin-top:10px; color:#808080;">
                                <i><strong>Prior: {poff_last:,}</strong></i>
                            </p>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
            st.markdown(f"""
                <hr style='margin-top: 3px; margin-bottom: 5px;'>
                <p style='font-size:10px; line-height:1.2; margin-top:0px; margin-bottom:1px;'>
                <strong>Decision Labels</strong>: {'; '.join(offered_keys)}
                </p>
            """, unsafe_allow_html=True)
        with top_intent.container(border=True, height=top_height):
            intent1, intent2 = st.columns([1.5, 1])
            intent1.markdown(f"""
                <p style='font-size:14px; line-height:0; margin-top:5px; text-align:left; color:#0D0D0D;'>
                <strong>Matriculation Intent</strong>
                </p>
                
                <p style='font-size:36px; line-height:1; margin-top:10px; text-align:left; color:#0D0D0D;'>
                {int_curr:,}
                </p>
                
                <div style='display: flex; justify-content: space-between; line-height:1; margin-top:-5px;'>
                    <span style='color:{int_pms[1]}; background-color:{int_pms[2]}; 
                                 padding:4px 6px; border-radius:10px; font-size:13px;'>
                    {int_pms[0]} {int_delta:,.0f} %
                    </span>
                    <span style='font-size:13px; background-color:#FAFAFA;
                                 padding:4px 8px; border-radius:5px; color:#808080;'>
                    <i><strong>Prior: {int_last:,}</strong></i>
                    </span>
                </div>
            """, unsafe_allow_html=True)
            with intent2:
                st.markdown(f"""
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div style="display:flex; flex-direction:column; align-items:flex-start; justify-content:space-between;">
                            <div style="font-size:11px; line-height:1; margin:0px; color:#0D0D0D;">
                                <strong>MS/ME</strong>
                            </div>
                            <div style="font-size:24px; line-height:1.2; color:#0D0D0D; ">
                                {mint_curr:,}
                            </div>
                        </div>
                        <div style="display:flex; flex-direction:column; align-items:flex-end; justify-content:space-between;">
                            <span style="font-size:11px; color:{mint_pms[1]}; background-color:{mint_pms[2]};
                                        padding:1px 4px; border-radius:10px; margin:-2px; ">
                                {mint_pms[0]} {mint_delta:,.0f} %
                            </span>
                            <p style="font-size:10px; line-height:1; margin-top:10px; color:#808080;">
                                <i><strong>Prior: {mint_last:,}</strong></i>
                            </p>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                st.markdown(f"""
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div style="display:flex; flex-direction:column; align-items:flex-start; justify-content:space-between;">
                            <div style="font-size:11px; line-height:1; margin:0px; color:#0D0D0D;">
                                <strong>PHD</strong>
                            </div>
                            <div style="font-size:24px; line-height:1.2; color:#0D0D0D; ">
                                {pint_curr:,}
                            </div>
                        </div>
                        <div style="display:flex; flex-direction:column; align-items:flex-end; justify-content:space-between;">
                            <span style="font-size:11px; color:{pint_pms[1]}; background-color:{pint_pms[2]};
                                        padding:1px 4px; border-radius:10px; margin:-2px;">
                                {pint_pms[0]} {pint_delta:,.0f} %
                            </span>
                            <p style="font-size:10px; line-height:1; margin-top:10px; color:#808080;">
                                <i><strong>Prior: {pint_last:,}</strong></i>
                            </p>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
            st.markdown(f"""
                <hr style='margin-top: 3px; margin-bottom: 5px;'>
                <p style='font-size:10px; line-height:1.2; margin-top:0px; margin-bottom:1px;'>
                <strong>Decision Labels</strong>: {'; '.join(intended_keys)}
                </p>
            """, unsafe_allow_html=True)
        with top_progress.container(border=False, height=top_height):
            value = int(int_curr / off_curr * 100) if off_curr else 0
            st.markdown(f"""
            <div style="display: flex; justify-content: center; align-items: center; height: {top_height*1.15}px; width: 100%;">
                <div style="position: relative; height: 100%; width: 100%; ">
                    <svg width="100%" height="100%" viewBox="0 0 140 140" preserveAspectRatio="xMidYMid meet"
                                style="position: relative; z-index: 0;">
                        <circle r="50" cx="60" cy="60" fill="transparent" stroke="#F2F2F2" stroke-width="15"/>
                        <circle r="50" cx="60" cy="60" fill="transparent" stroke="{convert_rgb(rgb_picked[1], 'hex')}"
                                stroke-width="15" stroke-dasharray="{value * 3.14}, 314" stroke-linecap="butt"
                                transform="rotate(-90 60 60)"/>
                    </svg>
                    <div style="position: absolute; top: 44%; left: 43%; transform: translate(-50%, -50%);
                                display: flex; align-items: center; justify-content: center; flex-direction: column;
                                font-weight: bold; z-index: 1; font-family: 'Verdana', sans-serif;
                                color: {convert_rgb(rgb_picked[1], 'hex')}">
                        <div style="font-size: 7px; line-height: 1;">Enrollment</div>
                        <div style="font-size: 7px; line-height: 1;">Progress</div>
                        <div style="font-size: 40px; line-height: 1;">{value}</div>
                        <div style="font-size: 8px; line-height: 1;">%</div>
                    </div>
                </div>
            </div>""", unsafe_allow_html=True)

    # MIDDLE: Tabs
    with st.container():
        tab1_row1_1, tab1_row1_2, tab1_row1_3 = st.columns([1, 1, 1.5])

        # Applicant Nationality Breakdown
        with tab1_row1_1.container(border=False, height=310):
            df_sunburst1 = df_submitted.groupby(["Degree", "Citizenship", "Continent"],
                                                dropna=False).size().reset_index(name="Value")
            df_sunburst1["Continent"] = np.where(df_sunburst1["Citizenship"] == "US", np.nan, df_sunburst1["Continent"])
            df_sunburst1["Citizenship"] = df_sunburst1["Citizenship"].replace({
                "FN": "Foreign", "US": "US Citizen", "PR": "Foreign", "": "Unknown"})
            c_labels = pd.unique(pd.concat([
                df_sunburst1['Degree'], df_sunburst1['Citizenship'], df_sunburst1["Continent"]], ignore_index=True))
            sunburst_color1_map = {c: rgb_picked[i] if i < len(rgb_picked) else "rgb(200,200,200)" for i, c in enumerate(c_labels)}
            try:
                sunburst1_fig = px.sunburst(
                    df_sunburst1,
                    path=["Degree", "Citizenship", "Continent"],
                    values="Value",
                    color="Value"
                )
                sunburst_labels1 = sunburst1_fig.data[0].labels
                colors_order1 = [sunburst_color1_map.get(lbl, "rgb(200,200,200)") for lbl in sunburst_labels1]
                sunburst1_fig.update_traces(marker=dict(colors=colors_order1))
                sunburst1_fig.update_layout(
                    height=275,
                    margin=dict(l=1, r=1, t=1, b=1),
                    coloraxis_showscale=False
                )
                sunburst1_fig.update_traces(
                    textinfo='label+percent parent',
                    insidetextorientation="horizontal",
                    maxdepth=4,
                    hovertemplate=''
                )
                st.plotly_chart(sunburst1_fig, use_container_width=True)
            except:
                pass
            st.markdown(f"""
                <p style='font-size:12px; line-height:0.5; margin-top:5px; margin-bottom:5px; text-align:center;'>
                <strong>Applicant Citizenship Breakdown</strong>
                </p>
            """, unsafe_allow_html=True)

        # Intended Matriculants Nationality Breakdown
        with tab1_row1_2.container(border=False, height=310):
            df_sunburst2 = df_intended.groupby(["Degree", "Citizenship", "Continent"],
                                               dropna=False).size().reset_index(name="Value")
            df_sunburst2["Continent"] = np.where(df_sunburst2["Citizenship"] == "US", np.nan, df_sunburst2["Continent"])
            df_sunburst2["Citizenship"] = df_sunburst2["Citizenship"].replace({
                "FN": "Foreign", "US": "US Citizen", "PR": "Foreign", "": "Unknown"})
            sunburst_color2_map = {c: rgb_picked[-i] if i < len(rgb_picked) else "rgb(200,200,200)" for i, c in
                                   enumerate(c_labels)}
            try:
                sunburst2_fig = px.sunburst(
                    df_sunburst2,
                    path=["Degree", "Citizenship", "Continent"],
                    values="Value",
                    color="Value"
                )
                sunburst2_labels = sunburst2_fig.data[0].labels
                colors2_order = [sunburst_color2_map.get(lbl, "rgb(200,200,200)") for lbl in sunburst2_labels]
                sunburst2_fig.update_traces(marker=dict(colors=colors2_order))
                sunburst2_fig.update_layout(
                    height=275,
                    margin=dict(l=1, r=1, t=1, b=1),
                    coloraxis_showscale=False
                )
                sunburst2_fig.update_traces(
                    textinfo='label+percent parent',
                    insidetextorientation="horizontal",
                    maxdepth=4,
                    hovertemplate=''
                )
                st.plotly_chart(sunburst2_fig, use_container_width=True)
            except:
                pass
            st.markdown(f"""
                <p style='font-size:12px; line-height:0.5; margin-top:5px; margin-bottom:5px; text-align:center;'>
                <strong>Intended Matriculants Citizenship Breakdown</strong>
                </p>
            """, unsafe_allow_html=True)

        # Submission Activity Bar Chart
        with tab1_row1_3.container(border=True, height=310):
            st.markdown(f"""
                <p style="font-size:14px; line-height:0.5; margin-top:5px; margin-bottom:15px; text-align:center;">
                <strong>Applicant Submission Activity for {curr_yr} Terms: Spring, Summer, Fall</strong>
                </p>
            """, unsafe_allow_html=True)
            df_tab1_bar = df_submitted.groupby(["Submission Month", "Continent"]).size().reset_index(
                name="Number of Applications")
            df_tab1_bar["Order"] = pd.to_datetime(df_tab1_bar["Submission Month"], format="%Y %b")
            bar_cutoff_early = pd.to_datetime(f"{last_yr} Aug", format="%Y %b")
            bar_cutoff_late = pd.to_datetime(f"{curr_yr} Mar", format="%Y %b")
            df_tab1_bar["Submission Month"] = np.where(
                df_tab1_bar["Order"] < bar_cutoff_early, "Before Aug",
                np.where(df_tab1_bar["Order"] > bar_cutoff_late, "After Mar",
                         df_tab1_bar["Submission Month"].str.slice(start=2)))
            continent_order = ["North America", "South America", "Asia", "Africa", "Europe", "Oceania"]
            bar_color_map = {c: convert_rgb(rgb_picked[i], "hex") for i, c in enumerate(continent_order)}
            df_tab1_bar.sort_values("Order", inplace=True)
            tab1_bar_fig = px.bar(
                df_tab1_bar,
                x="Submission Month",
                y="Number of Applications",
                color="Continent",
                color_discrete_map=bar_color_map,
                barmode="stack",
                category_orders={"Continent": continent_order}
            )
            tab1_bar_fig.update_layout(
                height=250,
                width=600,
                margin=dict(t=15, b=1, l=2, r=2),
                xaxis_title=None,
                yaxis_title=None,
                xaxis=dict(showgrid=True, tickfont=dict(size=10.5), tickangle=0),
                yaxis=dict(showgrid=True, range=[0, 800], tickfont=dict(size=11)),
                legend_title_text="",
                legend=dict(orientation="h", x=-0.05, y=-0.15, xanchor="left", yanchor="top", font=dict(size=9),
                            indentation=-8)
            )
            st.plotly_chart(tab1_bar_fig, use_container_width=True)

        st.markdown("<hr style='margin-top: 5px; margin-bottom: 5px;'>", unsafe_allow_html=True)
        st.markdown(f"""
            <p style="font-size:14px; line-height:0.5; margin-top:20px; margin-bottom:15px; text-align:left;">
            <strong>Applicant Background Analysis</strong>
            </p>
        """, unsafe_allow_html=True)
        tab1_row2_l1, tab1_row2_l2, tab1_row2_l3, tab1_row2_l4 = st.columns([0.5, 0.5, 1, 1])
        txt_c, bg_c = "black", "#F9FAFB"

        # Top 10 Countries
        with tab1_row2_l1.container(border=False, height=310):
            sub_tp10_cntry = get_top(
                df_submitted[df_submitted["Citizenship"].isin(["FN", "PR", "US"])]["Citizenship1"], "Country", 10)
            sub_cntry_width = sub_tp10_cntry["Percentage"].max()
            if sub_cntry_width:
                for row in range(sub_tp10_cntry.shape[0]):
                    sub_cntry = sub_tp10_cntry.loc[row, "Country"]
                    sub_cntry_ct = sub_tp10_cntry.loc[row, "Count"]
                    sub_cntry_pct = sub_tp10_cntry.loc[row, "Percentage"]
                    sub_cntry_margin = 0 if row < 9 else 10
                    st.markdown(f"""
                        <div style="width: 100%; height: 24px; margin-top: 4px; margin-bottom: {sub_cntry_margin}px;">
                            <div style="width: 100%; height: 95%; border-radius: 5px; color: {txt_c}; 
                                        background: {bg_c}; font-size: 12px; text-align: left; 
                                        padding-top: 2px; padding-bottom: 3px; padding-left: 5px;
                                        line-height: 20px; box-sizing: border-box;">
                                {sub_cntry} ({sub_cntry_ct}): {int(sub_cntry_pct) if int(sub_cntry_pct) > 0 else '<0'}%
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
            st.markdown(f"""
                <p style="font-size:12px; line-height:0.5; margin-top:5px; margin-bottom:5px; text-align:center;">
                <strong>Top 10 Nationalities</strong>
                </p>
            """, unsafe_allow_html=True)

        # Top 10 US States
        with tab1_row2_l2.container(border=False, height=310):
            sub_tp10_usstt = get_top(
                df_submitted[df_submitted["Active Country"] == "United States"]["Active Region"], "States", 10)
            sub_stt_width = sub_tp10_usstt["Percentage"].max()
            if sub_stt_width:
                for row in range(sub_tp10_usstt.shape[0]):
                    sub_usstt = sub_tp10_usstt.loc[row, "States"]
                    sub_usstt_ct = sub_tp10_usstt.loc[row, "Count"]
                    sub_usstt_pct = sub_tp10_usstt.loc[row, "Percentage"]
                    sub_usstt_margin = 0 if row < 9 else 10
                    st.markdown(f"""
                        <div style="width: 100%; height: 24px; margin-top: 4px; margin-bottom: {sub_usstt_margin}px;">
                            <div style="width: 100%; height: 95%; border-radius: 5px; color: {txt_c}; 
                                        background: {bg_c}; font-size: 12px; text-align: left; 
                                        padding-top: 2px; padding-bottom: 3px; padding-left: 5px;
                                        line-height: 20px; box-sizing: border-box;">
                                {sub_usstt} ({sub_usstt_ct}): {int(sub_usstt_pct)} %
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
            st.markdown(f"""
                <p style="font-size:12px; line-height:0.5; margin-top:5px; margin-bottom:5px; text-align:center;">
                <strong>Top 10 U.S. States</strong>
                </p>
            """, unsafe_allow_html=True)

        # Top 10 nonUS Undergrad
        with tab1_row2_l3.container(border=False, height=310):
            sub_tp10_fnuni = get_top(df_submitted[df_submitted["is_usuni"] == "no"]["School 1 Institution"],
                                     "Undergrad Institution", 10)
            sub_uni_width = sub_tp10_fnuni["Percentage"].max()
            if sub_uni_width:
                for row in range(sub_tp10_fnuni.shape[0]):
                    sub_fnuni = sub_tp10_fnuni.loc[row, "Undergrad Institution"]
                    sub_fnuni_ct = sub_tp10_fnuni.loc[row, "Count"]
                    sub_fnuni_pct = sub_tp10_fnuni.loc[row, "Percentage"]
                    sub_fnuni_margin = 0 if row < 9 else 10
                    st.markdown(f"""
                        <div style="width: 100%; height: 24px; margin-top: 4px; margin-bottom: {sub_fnuni_margin}px;">
                            <div style="width: 100%; height: 95%; border-radius: 5px; color: {txt_c}; 
                                        background: {bg_c}; font-size: 12px; text-align: left; 
                                        padding-top: 2px; padding-bottom: 3px; padding-left: 5px;
                                        line-height: 20px; box-sizing: border-box;">
                                {sub_fnuni} ({sub_fnuni_ct}): {int(sub_fnuni_pct)} %
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
            st.markdown(f"""
                <p style="font-size:12px; line-height:0.5; margin-top:5px; margin-bottom:5px; text-align:center;">
                <strong>Top 10 Non-U.S. Undergrad Institutions</strong>
                </p>
            """, unsafe_allow_html=True)

        # Top 10 US Undergrad
        with tab1_row2_l4.container(border=False, height=310):
            sub_tp10_usuni = get_top(df_submitted[df_submitted["is_usuni"] == "yes"]["School 1 Institution"],
                                     "Undergrad Institution", 10)
            sub_uni_width = sub_tp10_usuni["Percentage"].max()
            if sub_uni_width:
                for row in range(sub_tp10_usuni.shape[0]):
                    sub_usuni = sub_tp10_usuni.loc[row, "Undergrad Institution"]
                    sub_usuni_ct = sub_tp10_usuni.loc[row, "Count"]
                    sub_usuni_pct = sub_tp10_usuni.loc[row, "Percentage"]
                    sub_usuni_margin = 0 if row < 9 else 10
                    st.markdown(f"""
                        <div style="width: 100%; height: 24px; margin-top: 4px; margin-bottom: {sub_usuni_margin}px;">
                            <div style="width: 100%; height: 95%; border-radius: 5px; color: {txt_c}; 
                                        background: {bg_c}; font-size: 12px; text-align: left; 
                                        padding-top: 2px; padding-bottom: 3px; padding-left: 5px;
                                        line-height: 20px; box-sizing: border-box;">
                                {sub_usuni} ({sub_usuni_ct}): {int(sub_usuni_pct)} %
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
            st.markdown(f"""
                <p style="font-size:12px; line-height:0.5; margin-top:5px; margin-bottom:5px; text-align:center;">
                <strong>Top 10 U.S. Undergrad Institutions</strong>
                </p>
            """, unsafe_allow_html=True)

        st.markdown("<hr style='margin-top: 5px; margin-bottom: 5px;'>", unsafe_allow_html=True)
        st.markdown(f"""
            <p style="font-size:14px; line-height:0.5; margin-top:20px; margin-bottom:15px; text-align:left;">
            <strong>Intended Matriculants Background Analysis</strong>
            </p>
        """, unsafe_allow_html=True)
        tab1_row3_l1, tab1_row3_l2, tab1_row3_l3, tab1_row3_l4 = st.columns([0.5, 0.5, 1, 1])

        # Top 10 Countries
        with tab1_row3_l1.container(border=False, height=310):
            int_tp10_cntry = get_top(
                df_intended[df_intended["Citizenship"].isin(["FN", "PR", "US"])]["Citizenship1"], "Country", 10)
            int_cntry_width = int_tp10_cntry["Percentage"].max()
            if int_cntry_width:
                for row in range(int_tp10_cntry.shape[0]):
                    int_cntry = int_tp10_cntry.loc[row, "Country"]
                    int_cntry_ct = int_tp10_cntry.loc[row, "Count"]
                    int_cntry_pct = int_tp10_cntry.loc[row, "Percentage"]
                    int_cntry_margin = 0 if row < 9 else 10
                    st.markdown(f"""
                        <div style="width: 100%; height: 24px; margin-top: 4px; margin-bottom: {int_cntry_margin}px;">
                            <div style="width: 100%; height: 95%; border-radius: 5px; color: {txt_c}; 
                                        background: {bg_c}; font-size: 12px; text-align: left; 
                                        padding-top: 2px; padding-bottom: 3px; padding-left: 5px;
                                        line-height: 20px; box-sizing: border-box;">
                                {int_cntry} ({int_cntry_ct}): {int(int_cntry_pct) if int(int_cntry_pct) > 0 else '<0'}%
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
            st.markdown(f"""
                <p style="font-size:12px; line-height:0.5; margin-top:5px; margin-bottom:5px; text-align:center;">
                <strong>Top 10 Nationalities</strong>
                </p>
            """, unsafe_allow_html=True)

        # Top 10 US States
        with tab1_row3_l2.container(border=False, height=310):
            int_tp10_usstt = get_top(
                df_intended[df_intended["Active Country"] == "United States"]["Active Region"], "States", 10)
            int_stt_width = int_tp10_usstt["Percentage"].max()
            if int_stt_width:
                for row in range(int_tp10_usstt.shape[0]):
                    int_usstt = int_tp10_usstt.loc[row, "States"]
                    int_usstt_ct = int_tp10_usstt.loc[row, "Count"]
                    int_usstt_pct = int_tp10_usstt.loc[row, "Percentage"]
                    int_usstt_margin = 0 if row < 9 else 10
                    st.markdown(f"""
                        <div style="width: 100%; height: 24px; margin-top: 4px; margin-bottom: {int_usstt_margin}px;">
                            <div style="width: 100%; height: 95%; border-radius: 5px; color: {txt_c}; 
                                        background: {bg_c}; font-size: 12px; text-align: left; 
                                        padding-top: 2px; padding-bottom: 3px; padding-left: 5px;
                                        line-height: 20px; box-sizing: border-box;">
                                {int_usstt} ({int_usstt_ct}): {int(int_usstt_pct)} %
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
            st.markdown(f"""
                <p style="font-size:12px; line-height:0.5; margin-top:5px; margin-bottom:5px; text-align:center;">
                <strong>Top 10 U.S. States</strong>
                </p>
            """, unsafe_allow_html=True)

        # Top 10 nonUS Undergrad
        with tab1_row3_l3.container(border=False, height=310):
            int_tp10_fnuni = get_top(df_intended[df_intended["is_usuni"] == "no"]["School 1 Institution"],
                                     "Undergrad Institution", 10)
            int_uni_width = int_tp10_fnuni["Percentage"].max()
            if int_uni_width:
                for row in range(int_tp10_fnuni.shape[0]):
                    int_fnuni = int_tp10_fnuni.loc[row, "Undergrad Institution"]
                    int_fnuni_ct = int_tp10_fnuni.loc[row, "Count"]
                    int_fnuni_pct = int_tp10_fnuni.loc[row, "Percentage"]
                    int_fnuni_margin = 0 if row < 9 else 10
                    st.markdown(f"""
                        <div style="width: 100%; height: 24px; margin-top: 4px; margin-bottom: {int_fnuni_margin}px;">
                            <div style="width: 100%; height: 95%; border-radius: 5px; color: {txt_c}; 
                                        background: {bg_c}; font-size: 12px; text-align: left; 
                                        padding-top: 2px; padding-bottom: 3px; padding-left: 5px;
                                        line-height: 20px; box-sizing: border-box;">
                                {int_fnuni} ({int_fnuni_ct}): {int(int_fnuni_pct)} %
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
            st.markdown(f"""
                <p style="font-size:12px; line-height:0.5; margin-top:5px; margin-bottom:5px; text-align:center;">
                <strong>Top 10 Non-U.S. Undergrad  Institutions</strong>
                </p>
            """, unsafe_allow_html=True)

        # Top 10 US Undergrad
        with tab1_row3_l4.container(border=False, height=310):
            int_tp10_usuni = get_top(df_intended[df_intended["is_usuni"] == "yes"]["School 1 Institution"],
                                    "Undergrad Institution", 10)
            int_uni_width = int_tp10_usuni["Percentage"].max()
            if int_uni_width:
                for row in range(int_tp10_usuni.shape[0]):
                    int_usuni = int_tp10_usuni.loc[row, "Undergrad Institution"]
                    int_usuni_ct = int_tp10_usuni.loc[row, "Count"]
                    int_usuni_pct = int_tp10_usuni.loc[row, "Percentage"]
                    int_usuni_margin = 0 if row < 9 else 10
                    st.markdown(f"""
                        <div style="width: 100%; height: 24px; margin-top: 4px; margin-bottom: {int_usuni_margin}px;">
                            <div style="width: 100%; height: 95%; border-radius: 5px; color: {txt_c}; 
                                        background: {bg_c}; font-size: 12px; text-align: left; 
                                        padding-top: 2px; padding-bottom: 3px; padding-left: 5px;
                                        line-height: 20px; box-sizing: border-box;">
                                {int_usuni} ({int_usuni_ct}): {int(int_usuni_pct)} %
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
            st.markdown(f"""
                <p style="font-size:12px; line-height:0.5; margin-top:5px; margin-bottom:5px; text-align:center;">
                <strong>Top 10 U.S. Undergrad Institutions</strong>
                </p>
            """, unsafe_allow_html=True)

else:
    st.subheader("Summary Report")
    st.write("No Data Loaded")

