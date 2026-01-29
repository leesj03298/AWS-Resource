import boto3
import pandas as pd

# ==========================================
# 1. 설정 (Profile 및 기본 정보)
# ==========================================
AWS_PROFILE = "default"
REGION_NAME = "ap-northeast-2"
OUTPUT_FILE = "aws_vpce_with_eni_ip.xlsx"
# ==========================================

def get_vpce_data_with_ip():
    session = boto3.Session(profile_name=AWS_PROFILE, region_name=REGION_NAME)
    ec2 = session.client('ec2')

    # 이름 매핑용 마스터 데이터 수집
    vpcs = {v['VpcId']: next((t['Value'] for t in v.get('Tags', []) if t['Key'] == 'Name'), v['VpcId']) 
            for v in ec2.describe_vpcs()['Vpcs']}
    subnets = {s['SubnetId']: next((t['Value'] for t in s.get('Tags', []) if t['Key'] == 'Name'), s['SubnetId']) 
               for s in ec2.describe_subnets()['Subnets']}
    sgs = {g['GroupId']: next((t['Value'] for t in g.get('Tags', []) if t['Key'] == 'Name'), g['GroupId']) 
           for g in ec2.describe_security_groups()['SecurityGroups']}

    vpces = ec2.describe_vpc_endpoints()['VpcEndpoints']
    all_rows = []

    for idx, vpce in enumerate(vpces, 1):
        vpce_id = vpce['VpcEndpointId']
        vpce_type = vpce['VpcEndpointType']
        
        # 공통 기본 정보 (No, Name, Service, ID, Type, VPC, SG)
        base_info = {
            'No.': idx,
            'Name': next((t['Value'] for t in vpce.get('Tags', []) if t['Key'] == 'Name'), '-'),
            'Service Name': vpce['ServiceName'],
            'Endpoint ID': vpce_id,
            'Type': vpce_type,
            'VPC': vpcs.get(vpce['VpcId'], vpce['VpcId']),
            'Security Group': "\n".join([sgs.get(g['GroupId'], g['GroupId']) for g in vpce.get('Groups', [])]) or '-'
        }

        # Interface 타입: 서브넷별 ENI IP 추출
        if vpce_type == 'Interface' and vpce.get('NetworkInterfaceIds'):
            # ENI 상세 정보 조회 (Subnet ID와 Private IP 매핑 목적)
            eni_list = ec2.describe_network_interfaces(NetworkInterfaceIds=vpce['NetworkInterfaceIds'])['NetworkInterfaces']
            
            for eni in eni_list:
                row = base_info.copy()
                row.update({
                    'Subnet': subnets.get(eni['SubnetId'], eni['SubnetId']),
                    'Private IP': eni['PrivateIpAddress']
                })
                all_rows.append(row)
        
        # Gateway 타입: IP가 없으므로 하이픈 처리
        else:
            row = base_info.copy()
            row.update({'Subnet': '-', 'Private IP': '-'})
            all_rows.append(row)

    return pd.DataFrame(all_rows)

def save_with_styled_excel(df, filename):
    # 컬럼 순서 조정 (Subnet 뒤에 Private IP 배치)
    col_order = ['No.', 'Name', 'Service Name', 'Endpoint ID', 'Type', 'VPC', 'Subnet', 'Private IP', 'Security Group']
    df = df[col_order]

    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='VPCEndpoints')
    workbook, worksheet = writer.book, writer.sheets['VPCEndpoints']
    
    # 공통 스타일 (가운데 정렬 + 테두리 + 텍스트 줄바꿈)
    fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center', 'valign': 'vcenter'})

    # 헤더 적용
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)

    # 병합 로직 (Endpoint ID 기준)
    # 병합할 열: No(0), Name(1), Service(2), ID(3), Type(4), VPC(5), Security Group(8)
    merge_cols = [0, 1, 2, 3, 4, 5, 8]
    
    for col in merge_cols:
        start_row = 1
        while start_row <= len(df):
            end_row = start_row
            # Endpoint ID(3번 컬럼)가 같으면 병합 대상으로 판단
            while end_row < len(df) and df.iloc[end_row, 3] == df.iloc[start_row - 1, 3]:
                end_row += 1
            
            val = df.iloc[start_row - 1, col]
            if end_row - start_row > 0:
                worksheet.merge_range(start_row, col, end_row, col, val, fmt)
            else:
                worksheet.write(start_row, col, val, fmt)
            start_row = end_row + 1

    # Subnet(6) 및 Private IP(7) 열 개별 스타일 적용
    for r_idx in range(1, len(df) + 1):
        worksheet.write(r_idx, 6, df.iloc[r_idx-1, 6], fmt)
        worksheet.write(r_idx, 7, df.iloc[r_idx-1, 7], fmt)

    # 열 너비 최적화
    worksheet.set_column('A:A', 6)   # No.
    worksheet.set_column('B:D', 35)  # Name, Service, ID
    worksheet.set_column('E:F', 15)  # Type, VPC
    worksheet.set_column('G:I', 25)  # Subnet, IP, SG
    
    writer.close()
    print(f"✨ 추출 완료! 파일명: {filename}")

if __name__ == "__main__":
    df_result = get_vpce_data_with_ip()
    save_with_styled_excel(df_result, OUTPUT_FILE)