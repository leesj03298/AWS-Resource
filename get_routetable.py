import boto3
import pandas as pd

# ==========================================
# 1. 설정 (Profile 및 기본 정보)
# ==========================================
AWS_PROFILE = "default"
REGION_NAME = "ap-northeast-2"
ACCOUNT_LABEL = "DEV"
OUTPUT_FILE = "aws_route_table_final.xlsx"
# ==========================================

def get_full_data():
    session = boto3.Session(profile_name=AWS_PROFILE, region_name=REGION_NAME)
    ec2 = session.client('ec2')
    
    # VPC 정보 수집
    vpcs = ec2.describe_vpcs()['Vpcs']
    vpc_map = {v['VpcId']: next((t['Value'] for t in v.get('Tags', []) if t['Key'] == 'Name'), v['VpcId']) for v in vpcs}
    
    # Route Table 정보 수집
    rtbs = ec2.describe_route_tables()['RouteTables']
    
    rows = []
    for rtb in rtbs:
        vpc_id = rtb['VpcId']
        vpc_name = vpc_map.get(vpc_id, 'N/A')
        rtb_id = rtb['RouteTableId']
        rtb_name = next((tag['Value'] for tag in rtb.get('Tags', []) if tag['Key'] == 'Name'), 'Unused')
        
        for route in rtb.get('Routes', []):
            dest = route.get('DestinationCidrBlock') or route.get('DestinationPrefixListId') or '-'
            target = route.get('GatewayId') or route.get('TransitGatewayId') or \
                     route.get('NatGatewayId') or route.get('NetworkInterfaceId') or \
                     route.get('VpcPeeringConnectionId') or '-'
            
            if 'GatewayId' in route and route['GatewayId'] == 'local':
                target = 'local'

            rows.append({
                'ACCOUNT': ACCOUNT_LABEL,
                'VPC Name': vpc_name,
                'VPC ID': vpc_id,
                'Route Tables Name': rtb_name,
                'Route Tables ID': rtb_id,
                'Destination': dest,
                'Target': target
            })
    
    df = pd.DataFrame(rows)
    # 정렬 가중치: local을 0순위로
    df['Target_Priority'] = df['Target'].apply(lambda x: 0 if x == 'local' else 1)
    df = df.sort_values(by=['VPC Name', 'Route Tables Name', 'Target_Priority', 'Destination'])
    return df.drop(columns=['Target_Priority'])

def save_with_merging_centered(df, filename):
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='RouteTables')
    
    workbook = writer.book
    worksheet = writer.sheets['RouteTables']
    
    # ---------------------------------------------------------
    # 공통 스타일 정의 (가운데 정렬 추가)
    # ---------------------------------------------------------
    # 1. 헤더 스타일: 굵게, 배경색, 테두리, 가로/세로 가운데 정렬
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 
        'align': 'center', 'valign': 'vcenter'
    })
    
    # 2. 병합 및 일반 셀 스타일: 테두리, 가로/세로 가운데 정렬
    center_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })

    # 헤더 적용
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # 병합 로직 (0~4번 컬럼: ACCOUNT, VPC Name, VPC ID, RT Name, RT ID)
    merge_cols = [0, 1, 2, 3, 4]
    for col in merge_cols:
        start_row = 1
        while start_row <= len(df):
            end_row = start_row
            current_val = df.iloc[start_row - 1, col]
            
            while end_row < len(df) and df.iloc[end_row, col] == current_val:
                end_row += 1
            
            if end_row - start_row > 0:
                worksheet.merge_range(start_row, col, end_row, col, current_val, center_format)
            else:
                worksheet.write(start_row, col, current_val, center_format)
            start_row = end_row + 1

    # Destination(5) / Target(6) 열에도 가운데 정렬 스타일 적용
    for r_idx in range(1, len(df) + 1):
        worksheet.write(r_idx, 5, df.iloc[r_idx-1, 5], center_format)
        worksheet.write(r_idx, 6, df.iloc[r_idx-1, 6], center_format)

    # 열 너비 조정
    worksheet.set_column('A:G', 25)
    
    writer.close()
    print(f"✨ 모든 셀 가운데 정렬 완료: {filename}")

if __name__ == "__main__":
    final_df = get_full_data()
    save_with_merging_centered(final_df, OUTPUT_FILE)