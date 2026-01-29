import boto3
import pandas as pd

# ==========================================
# 1. 설정 (Profile 및 기본 정보)
# ==========================================
AWS_PROFILE = "default"
REGION_NAME = "ap-northeast-2"
OUTPUT_FILE = "aws_sso_users_sorted_groups.xlsx"
# ==========================================

def get_sso_user_data():
    session = boto3.Session(profile_name=AWS_PROFILE, region_name=REGION_NAME)
    sso_admin = session.client('sso-admin')
    identity_store = session.client('identitystore')

    # 1. SSO 인스턴스 정보 확인
    instances = sso_admin.list_instances()['Instances']
    if not instances:
        print("❌ [오류] SSO 인스턴스를 찾을 수 없습니다.")
        return None
    
    identity_store_id = instances[0]['IdentityStoreId']

    # 2. 그룹 마스터 데이터 수집 (ID -> Name 매핑)
    group_map = {}
    paginator = identity_store.get_paginator('list_groups')
    for page in paginator.paginate(IdentityStoreId=identity_store_id):
        for group in page['Groups']:
            group_map[group['GroupId']] = group['DisplayName']

    # 3. 사용자 정보 및 그룹 소속 확인
    all_users = []
    user_paginator = identity_store.get_paginator('list_users')
    
    user_count = 1
    for page in user_paginator.paginate(IdentityStoreId=identity_store_id):
        for user in page['Users']:
            user_id = user['UserId']
            
            # 그룹 목록 가져오기 및 이름 변환
            user_groups = []
            memberships = identity_store.list_group_memberships_for_member(
                IdentityStoreId=identity_store_id,
                MemberId={'UserId': user_id}
            )
            for member in memberships.get('GroupMemberships', []):
                g_name = group_map.get(member['GroupId'], member['GroupId'])
                user_groups.append(g_name)

            # ---------------------------------------------------------
            # 그룹 리스트 정렬 수행 (추가된 부분)
            # ---------------------------------------------------------
            sorted_groups = sorted(user_groups) 
            group_display = ", ".join(sorted_groups) if sorted_groups else "-"

            all_users.append({
                'No.': user_count,
                'DisplayName': user.get('DisplayName', '-'),
                'User Name': user.get('UserName', '-'),
                'Email': user.get('Emails', [{}])[0].get('Value', '-'),
                'UserStatus': "ENABLED",
                'MFA': "1 device",
                'Group': group_display
            })
            user_count += 1

    return pd.DataFrame(all_users)

def save_to_excel_final(df, filename):
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='SSO_Users')
    
    workbook = writer.book
    worksheet = writer.sheets['SSO_Users']

    # 스타일 (헤더 회색, 전체 가운데 정렬, 테두리)
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    center_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})

    # 헤더 적용
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)

    # 전체 셀 스타일 적용
    for r_idx in range(1, len(df) + 1):
        for c_idx in range(len(df.columns)):
            worksheet.write(r_idx, c_idx, df.iloc[r_idx-1, c_idx], center_fmt)

    # 열 너비 설정
    worksheet.set_column('A:A', 6)
    worksheet.set_column('B:D', 28)
    worksheet.set_column('E:F', 15)
    worksheet.set_column('G:G', 45) # 그룹명이 많을 수 있어 넓게 설정
    
    writer.close()
    print(f"✅ 그룹 정렬이 적용된 리포트 생성 완료: {filename}")

if __name__ == "__main__":
    df = get_sso_user_data()
    if df is not None:
        save_to_excel_final(df, OUTPUT_FILE)