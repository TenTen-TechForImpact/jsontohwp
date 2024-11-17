import win32com.client as win32
import os
from datetime import datetime

def format_date(date_str):
    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        return date_obj.strftime(" %Y  /   %m   /  %d")
    except ValueError as ve:
        print(f"날짜 포맷 오류: {ve}")
        return ""
    except TypeError as te:
        print(f"날짜 값이 비어 있습니다: {te}")
        return ""

#insert text
def set_text(hwp, field_name, value):
    try:
        hwp.PutFieldText(field_name,value)
    except Exception as e:
        print(f"필드 '{field_name}' 설정 오류: {e}")

#insert list
def insert_list_into_table(hwp, field, data_list, separate=True):
    """
    separate=True: list의 item 다른 셀에 삽입
    separate=False: list의 item 모두 한 셀에 삽입
    """
    try:
        if separate:
            set_text(hwp, field, data_list[0])
            for idx, item in enumerate(data_list, start=1):
                hwp.Run("TableBelowCell")
                hwp.Run("TableSelectCell")
                hwp.HAction.Run("InsertText", f"{idx}. {item}")
        else:
            combined_text = ", ".join(data_list)
            set_text(hwp, field, combined_text)
    except Exception as e:
        print(f"리스트 데이터 삽입 오류: {e}")

#checkbox
def set_checkbox(hwp, field_yes, field_no, value='아니오'):
    """
    :param value: '예' 또는 '아니오'
    """
    try:
        if value == "예":
            hwp.PutFieldText(field_yes, "\u25A0")
            hwp.PutFieldText(field_no, "\u25A1")
        elif value == "아니오":
            hwp.PutFieldText(field_yes, "\u25A1")
            hwp.PutFieldText(field_no, "\u25A0")
        else:
            print(f"체크박스 오류: {e}")
    except Exception as e:
        print(f"체크박스 오류: {e}")


#전문의약품
def insert_edrugs(hwp, ethical_drugs):
    try:
        for idx, med in enumerate(ethical_drugs, start=1):
            set_text(hwp, f"{idx}_NAME", med.get('name', ''))
            set_text(hwp, f"{idx}_DAYS", str(med.get('prescription_days', '')))
            set_text(hwp, f"{idx}_PURPOSE", med.get('purpose', ''))
            set_text(hwp, f"{idx}_USAGE_STATUS", med.get('usage_status', ''))
            
    except Exception as e:
        print(f"ethical_the_counter_drugs 데이터 삽입 오류: {e}")

#건강의약품+일반의약품
def insert_odrugs_healthfood(hwp,over_drugs, health_foods):
    try:
        combined_medications = []

        for med in over_drugs:
            combined_medications.append({
                'name': med.get('name', ''),
                'unit': med.get('unit', ''),
                'purpose': med.get('purpose', ''),
                'usage_status': med.get('usage_status', ''),
                'type': '일반의약품'
            })

        for food in health_foods:
            combined_medications.append({
                'name': food.get('name', ''),
                'unit': food.get('unit', ''),
                'purpose': food.get('purpose', ''),
                'usage_status': food.get('usage_status', ''),
                'type': '건강기능식품'
            })
        
        for idx, med in enumerate(combined_medications, start=1):
            fieldnames={
                'type': f"{idx}_TYPE",
                'name': f"{idx}_NAMEs",
                'unit': f"{idx}_UNIT",
                'purpose': f"{idx}_PURPOSEs",
                'usage_status': f"{idx}_USAGE_STATUSs"
            }
            set_text(hwp, fieldnames['type'], med['type'])
            set_text(hwp, fieldnames['name'], med['name'])
            set_text(hwp, fieldnames['unit'], med['unit'])
            set_text(hwp, fieldnames['purpose'], med['purpose'])
            set_text(hwp, fieldnames['usage_status'], med['usage_status'])
    except Exception as e:
        print(f"약품 데이터 삽입 오류: {e}")

def find_matching_field(disease, disease_fields):
    # 키에 부분적으로 일치
    for option, (field_yes, field_no) in disease_fields.items():
        if disease in option:
            return field_yes, field_no
    return None

def create_hwp_file(json_data):

    # HWP 오토메이션 초기화
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.XHwpWindows.Item(0).Visible = False
    except Exception as e:
        print(f"HWP 오토메이션 초기화 오류: {e}")
        return None

    # 템플릿 파일 열기
    template_path = './files/template.hwp'
    output_dir = './files'
    output_file = 'output.hwp'
    output_path = os.path.join(output_dir, output_file)

    template_path_abs = os.path.abspath(template_path)
    output_path_abs = os.path.abspath(output_path)

    try:
        hwp.Open(template_path_abs, "HWP", None)
    except Exception as e:
        print(f"HWP 템플릿 열기 오류: {e}")
        return None

    # personal_info
    personal_info = json_data.get('personal_info', {})
    set_text(hwp, "NAME", personal_info.get('name', ''))
    set_text(hwp, "DOB", personal_info.get('date_of_birth', ''))
    set_text(hwp, "PHONE", personal_info.get('phone_number', ''))

    consultation_info = json_data.get('consultation_info', {})
    initial_consult_date=consultation_info.get('initial_consult_date', '')
    converted=format_date(initial_consult_date)
    set_text(hwp, "INITIAL_CONSULT_DATE", converted)

    current_consult_date=consultation_info.get('current_consult_date', '')
    converted=format_date(current_consult_date)
    set_text(hwp, "CURRENT_CONSULT_DATE", converted)

    # 상담 약사
    pharmacist_names = consultation_info.get('pharmacist_names', [])
    pharmacist_name = (pharmacist_names + [""] *3)[:3]
    set_text(hwp, "PHARM1", pharmacist_names[0])
    set_text(hwp, "PHARM2", pharmacist_names[1])
    set_text(hwp, "PHARM3", pharmacist_names[2])

    # consultation_info
    # consultation_info = data.get('consultation_info', {})
    consultation_info = json_data.get('consultation_info', {})
    insurance_type = consultation_info.get('insurance_type', '')
    insurance_field = {
        "건강보험": ("INSURANCE_HEALTH_YES","INSURANCE_HEALTH_NO"),
        "의료급여": ("INSURANCE_MEDICAL_YES","INSURANCE_MEDICAL_NO"),
        "보훈": ("INSURANCE_BOHUN_YES","INSURANCE_BOHUN_NO"),
        "비급여": ("INSURANCE_NONPAY_YES","INSURANCE_NONPAY_NO")
    }
    for type in insurance_field:
        if type == insurance_type:
            field = insurance_field.get(type)
            set_checkbox(hwp, field[0], field[1], "예")
        else:
            field = insurance_field.get(type)
            set_checkbox(hwp, field[0], field[1], "아니오")
     
    # 만성질환 체크박스
    medical_conditions = json_data.get('medical_conditions', {})
    chronic_diseases = medical_conditions.get('chronic_diseases', {}).get('disease_names', [])
    disease_fields = {
        "고혈압": ("DISEASE_HYPERTENSION_YES", "DISEASE_HYPERTENSION_NO"),
        "고지혈증": ("DISEASE_HYPERLIPIDEMIA_YES", "DISEASE_HYPERLIPIDEMIA_NO"),
        "뇌혈관질환": ("DISEASE_CEREBROVASCULAR_YES", "DISEASE_CEREBROVASCULAR_NO"),
        "심장질환": ("DISEASE_HEART_YES", "DISEASE_HEART_NO"),
        "당뇨병": ("DISEASE_DIABETES_YES", "DISEASE_DIABETES_NO"),
        "갑상선질환": ("DISEASE_THYROID_YES", "DISEASE_THYROID_NO"),
        "위장관질환": ("DISEASE_GASTROINTESTINAL_YES", "DISEASE_GASTROINTESTINAL_NO"),
        "파킨슨": ("DISEASE_PARKINSON_YES", "DISEASE_PARKINSON_NO"),
        "척추·관절염/신경통·근육통": ("DISEASE_SPINE_YES", "DISEASE_SPINE_NO"),
        "수면장애": ("DISEASE_SLEEP_YES", "DISEASE_SLEEP_NO"),
        "우울증/불안장애": ("DISEASE_DEPRESSION_YES", "DISEASE_DEPRESSION_NO"),
        "치매,인지장애": ("DISEASE_DEMENTIA_YES", "DISEASE_DEMENTIA_NO"),
        "비뇨·생식기질환(전립선비대증,자궁내막염,방광염 등)": ("DISEASE_GENITOURINARY_YES", "DISEASE_GENITOURINARY_NO"),
        "신장질환": ("DISEASE_KIDNEY_YES", "DISEASE_KIDNEY_NO"),
        "호흡기질환(천식,COPD 등)": ("DISEASE_RESPIRATORY_YES", "DISEASE_RESPIRATORY_NO"),
        "안질환(백내장,녹내장,안구건조증 등)": ("DISEASE_OCULAR_YES", "DISEASE_OCULAR_NO"),
        "이비인후과(만성비염, 만성중이염 등)": ("DISEASE_OTOLARYNGOLOGY_YES", "DISEASE_OTOLARYNGOLOGY_NO"),
        "암질환": ("DISEASE_CANCER_YES", "DISEASE_CANCER_NO"),
        "간질환": ("DISEASE_LIVER_YES", "DISEASE_LIVER_NO"),
        "뇌경색": ("DISEASE_STROKE_YES", "DISEASE_STROKE_NO")
    }
    for option, (field_yes, field_no) in disease_fields.items():
        set_checkbox(hwp, field_yes, field_no, "아니오")
    for disease in chronic_diseases:
        field = find_matching_field(disease, disease_fields)
        if field:
            set_checkbox(hwp, field[0], field[1], "예")
    # 기타
    set_text(hwp, "ADDITIONAL_INFO", medical_conditions.get('chronic_diseases', {}).get('additional_info', ''))
    #과거 질병 및 수술 이력
    set_text(hwp, "MEDICAL_HISTORY", medical_conditions.get('medical_history', ''))
    #주요 불편한 증상
    set_text(hwp, "SYMPTOMS", medical_conditions.get('symptoms', ''))
 
    
    #알레르기 또는 약물부작용 확인
    # 알러지 체크박스
    has_allergies = medical_conditions.get('allergies', {}).get('has_allergies', '아니오')
    set_checkbox(hwp, "HAS_ALLERGIES_YES", "HAS_ALLERGIES_NO", has_allergies)
    allergies_suspected_item = medical_conditions.get('allergies', {}).get('suspected_items', [])
    insert_list_into_table(hwp, "SUSPECTED_ITEMS", allergies_suspected_item, False)
    # 약물부작용 체크박스
    has_adverse_drug_reactions = medical_conditions.get('adverse_drug_reactions', {}).get('has_adverse_drug_reactions', '아니오')
    set_checkbox(hwp, "HAS_ADVERSE_DRUG_REACTIONS_YES", "HAS_ADVERSE_DRUG_REACTIONS_NO", has_adverse_drug_reactions)
    suspected_medications = medical_conditions.get('adverse_drug_reactions', {}).get('suspected_medications', [])
    insert_list_into_table(hwp, "SUSPECTED_MEDICATIONS", suspected_medications, False)
    reaction_details = medical_conditions.get('adverse_drug_reactions', {}).get('reaction_details', [])
    insert_list_into_table(hwp, "REACTION_DETAILS", reaction_details, False)

    #생활습관
    # 흡연
    lifestyle = json_data.get('lifestyle', {})
    is_smoking = lifestyle.get('smoking', {}).get('is_smoking', '아니오')
    set_checkbox(hwp, "IS_SMOKING_YES", "IS_SMOKING_NO", is_smoking)
    if is_smoking == "예":
        set_text(hwp, "SMOKING_DURATION", str(lifestyle.get('smoking', {}).get('duration_in_years', '')))
        set_text(hwp, "PACK_PER_DAY", str(lifestyle.get('smoking', {}).get('pack_per_day', '')))
    #음주
    is_drinking = lifestyle.get('alcohol', {}).get('is_drinking', '아니오')
    set_checkbox(hwp, "IS_DRINKING_YES", "IS_DRINKING_NO", is_drinking)
    if is_drinking == "예":
        set_text(hwp, "DRINKS_PER_WEEK", str(lifestyle.get('alcohol', {}).get('drinks_per_week','')))
        set_text(hwp, "AMOUNT_PER_WEEK", lifestyle.get('alcohol',{}).get('amount_per_week', ''))
    #운동
    is_exercising = lifestyle.get('exercise', {}).get('is_exercising', '아니오')
    set_checkbox(hwp, "IS_EXERCISING_YES", "IS_EXERCISING_NO", is_exercising)
    if is_exercising == "예":
        exercise_frequency_map={
            "주1회": ("EXERCISE_FREQ_1_YES", "EXERCISE_FREQ_1_NO"),
            "주2회": ("EXERCISE_FREQ_2_YES", "EXERCISE_FREQ_2_NO"),
            "주3회": ("EXERCISE_FREQ_3_YES", "EXERCISE_FREQ_3_NO"),
            "주4회이상": ("EXERCISE_FREQ_4_YES", "EXERCISE_FREQ_4_NO")
        }
        exercise_frequency = lifestyle.get('exercise', {}).get('exercise_frequency', '')
        for fields in exercise_frequency_map.values():
            set_checkbox(hwp, fields[0], fields[1], "아니오")
        selected_fields = exercise_frequency_map.get(exercise_frequency)
        if selected_fields:
            set_checkbox(hwp, selected_fields[0], selected_fields[1], "예")
    exercise_types = lifestyle.get('exercise', {}).get('exercise_types', [])
    insert_list_into_table(hwp, "EXERCISE_TYPES", exercise_types, separate=False)
    #영양상태
    is_balanced_meal = lifestyle.get('diet', {}).get('is_balanced_meal', '아니오')
    set_checkbox(hwp, "IS_BALANCED_MEAL_YES", "IS_BALANCED_MEAL_NO", is_balanced_meal)
    if is_balanced_meal == "예":
        meal_map={
            "1회": ("MEAL_FREQ_1_YES", "MEAL_FREQ_1_NO"),
            "2회": ("MEAL_FREQ_2_YES", "MEAL_FREQ_2_NO"),
            "3회": ("MEAL_FREQ_3_YES", "MEAL_FREQ_3_NO"),
        }
        balanced_meals_per_day = lifestyle.get('diet', {}).get('balanced_meals_per_day', '')
        for fields in meal_map.values():
            set_checkbox(hwp, fields[0], fields[1], "아니오")
        selected_fields = meal_map.get(balanced_meals_per_day)
        if selected_fields:
            set_checkbox(hwp, selected_fields[0], selected_fields[1], "예")
    
    #약 복용 관리
    medication_management = json_data.get('medication_management',{})
    #독거여부
    living_alone = medication_management.get('living_condition', {}).get('living_alone', '예')
    set_checkbox(hwp, "LIVING_ALONE_YES", "LIVING_ALONE_NO", living_alone)
    if living_alone == "아니오":
        insert_list_into_table(hwp, "FAMILY_MEMBERS", medication_management.get('living_condition', {}).get('family_members',[]), False)
    #투약보조자
    medication_assistants = medication_management.get('living_condition', {}).get('medication_assistants', [])
    assistants_field = {
        "본인": ("ASSIST_SELF_YES", "ASSIST_SELF_NO"),
        "배우자": ("ASSIST_SPOUSE_YES", "ASSIST_SPOUSE_NO"),
        "자녀": ("ASSIST_CHILD_YES", "ASSIST_CHILD_NO"),
        "친인척": ("ASSIST_RELATIVE_YES", "ASSIST_RELATIVE_NO"), 
        "친구": ("ASSIST_FRIEND_YES", "ASSIST_FRIEND_NO"), 
        "요양보호사 또는 돌봄종사자": ("ASSIST_CAREGIVER_YES", "ASSIST_CAREGIVER_NO")
    }
    other_assistants = []
    for option, (field_yes, field_no) in assistants_field.items():
        set_checkbox(hwp, field_yes, field_no, "아니오")
    for assistant in medication_assistants:
        if assistant in assistants_field:
            field_yes, field_no = assistants_field[assistant]
            set_checkbox(hwp, field_yes, field_no, "예")
        elif find_matching_field(assistant, "요양보호사 또는 돌봄종사자"):
            field=assistants_field.get("요양보호사 또는 돌봄종사자")
            set_checkbox(hwp, field[0], field[1], "예")
        else:
            other_assistants.append(assistant)
            set_checkbox(hwp, "OTHERS_YES", "OTHERS_NO", "예")
    if other_assistants:
        other_text = ", ".join(other_assistants)
        set_text(hwp, "ASSIST_OTHER_FIELD", other_text)
    else:
        set_text(hwp, "ASSIST_OTHER_FIELD", "")

    #약 보관 장소
    has_medication_storage = medication_management.get('medication_storage', {}).get('has_medication_storage','아니오')
    set_checkbox(hwp, "HAS_MEDICATION_STORAGE_YES", "HAS_MEDICATION_STORAGE_NO", has_medication_storage)
    if has_medication_storage == "예":
        set_text(hwp, "LOCATION", medication_management.get('medication_storage', {}).get('location', ''))

    #처방전 보관여부
    is_prescription_stored = medication_management.get('prescription_storage', {}).get('is_prescription_stored','아니오')
    set_checkbox(hwp, "IS_PRESCRIPTION_STORED_YES", "IS_PRESCRIPTION_STORED_NO", is_prescription_stored)

    #복용 약물 개수
    current_medications = json_data.get('current_medications', {})
    edrugcount = current_medications.get('ethical_the_counter_drugs', {}).get('count', '')
    odrugcount = current_medications.get('over_the_counter_drugs', {}).get('count', '')
    healthfoodcount = current_medications.get('health_functional_foods', {}).get('count', '')
    set_text(hwp, "EDRUG_COUNT", edrugcount)
    set_text(hwp, "ODRUG_COUNT", odrugcount)
    set_text(hwp, "HEALTH_FUNCTIONAL_FOOD_COUNT", healthfoodcount)

    #current_medications
    insert_edrugs(hwp, current_medications.get('ethical_the_counter_drugs', {}).get('list', []))
    insert_odrugs_healthfood(hwp, current_medications.get('over_the_counter_drugs', {}).get('list', []), current_medications.get('health_functional_foods', {}).get('list', []))
    

    questions = json_data.get('questions', {}).get('list', [])
    questions_text = "\r\n".join([f"{idx+1}. {q}" for idx, q in enumerate(questions)])
    set_text(hwp, "QUESTIONS", questions_text)

    pharmacist_comments = json_data.get('pharmacist_comments', '')
    care_note = json_data.get('care_note', '')
    prefix=consultation_info.get('consult_session_number', '')
    set_text(hwp, f"{prefix}PHARMACIST_COMMENTS", pharmacist_comments)
    set_text(hwp, f"{prefix}CARE_NOTE", care_note)
    set_text(hwp, "1CONSLTDATE", initial_consult_date)
    set_text(hwp, f"{prefix}CONSLTDATE", current_consult_date)
    set_text(hwp, "1CONSLTPHARM", pharmacist_names[0])
    set_text(hwp, "2CONSLTPHARM", pharmacist_names[1])
    set_text(hwp, "3CONSLTPHARM", pharmacist_names[2])

    # HWP 파일 저장
    try:
        hwp.SaveAs(output_path_abs)
        hwp.Quit()
        print(f"데이터가 삽입된 HWP 파일이 '{output_path}'에 저장되었습니다.")
    except Exception as e:
        print(f"HWP 파일 저장 오류: {e}")
        return None

    # HWP 파일을 바이너리 모드로 읽어서 데이터 반환
    try:
        with open(output_path_abs, 'rb') as f:
            hwp_content = f.read()
        return hwp_content
    except Exception as e:
        print(f"HWP 파일 읽기 오류: {e}")
        return None

    except Exception as e:
        print(f"HWP 파일 생성 중 오류 발생: {e}")
        return None