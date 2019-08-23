import json
import io
import xlrd
import pandas as pd

def create_ec2():
    cfnDict = {
        "AWSTemplateFormatVersion": "2010-09-09",
        "Resources": {
            "MyInstance": {
                "Type": "AWS::EC2::Instance",
                "Properties": {}
            }
        }
    }

    loc = ("C:\\Users\\mayank\\Documents\\pythonWorkspace\\autogenerateCFN\\EC2Details.xlsx")

    workbook = xlrd.open_workbook(loc, on_demand=True)
    worksheet = workbook.sheet_by_index(0)
    first_row = []  # The row where we stock the name of the column
    for col in range(worksheet.ncols):
        first_row.append(worksheet.cell_value(0, col))
    data = {}
    for row in range(1, 2):
        elm = {}
        for col in range(worksheet.ncols):
            if worksheet.cell_value(row, col) != '' and first_row[col] != 'BlockDeviceMappingsDeviceName' and first_row[
                col] != 'BlockDeviceMappingsEBSVolumeSize' \
                    and first_row[col] != 'ElasticGpuSpecifications' \
                    and first_row[col] != 'ElasticInferenceAccelerators' and first_row[col] != 'Ipv6AddressCount' and \
                    first_row[col] != 'LicenseSpecifications' \
                    and first_row[col] != 'NetworkInterfaces' and first_row[col] != 'SecurityGroupIds' and first_row[
                col] != 'SecurityGroups' \
                    and first_row[col] != 'SsmAssociations' and first_row[col] != 'Tags' and first_row[
                col] != 'Volumes' and first_row[col] != 'TagKey' \
                    and first_row[col] != 'TagValue':
                elm[first_row[col]] = worksheet.cell_value(row, col)
        data = elm
    ##print (json.dumps(data,indent=4))

    Properties = {
        'BlockDeviceMappings': [],
        'ElasticGpuSpecifications': [],
        'ElasticInferenceAccelerators': [],
        'Ipv6Addresses': [],
        'LicenseSpecifications': [],
        'NetworkInterfaces': [],
        'SecurityGroupIds': [],
        'SecurityGroups': [],
        'SsmAssociations': [],
        'Tags': [],
        'Volumes': [],
    }

    BDMapping = []
    for i in range(1, worksheet.nrows):
        if worksheet.cell_value(i, 2) != '' and worksheet.cell_value(i, 3) != '':
            dv = worksheet.cell_value(i, 2)
            vs = (worksheet.cell_value(i, 3))
            BDMappingdict = {
                "DeviceName": dv,
                "Ebs": {
                    "VolumeSize": str(int(vs))
                }
            }
            BDMapping.append(BDMappingdict)
    ##print(BDMapping)

    taglist = []
    for j in range(2, worksheet.nrows):
        if worksheet.cell_value(j, 30) != '' and worksheet.cell_value(j, 31) != '':
            tdict = {}
            tdict['Key'] = worksheet.cell_value(j, 30)
            tdict['Value'] = worksheet.cell_value(j, 31)
            # print(tdict)
            taglist.append(tdict)
    ##print(taglist)
    df = pd.read_excel("C:\\Users\\mayank\\Documents\\pythonWorkspace\\autogenerateCFN\\EC2Details.xlsx",
                       sheet_name='EC2Details')  # can also index sheet by name or fetch all sheets

    ElasticGpuSpecificationsList = df['ElasticGpuSpecifications'].tolist()
    ElasticGpuSpecificationsList = [x for x in ElasticGpuSpecificationsList if str(x) != 'nan']

    ElasticInferenceAcceleratorsList = df['ElasticInferenceAccelerators'].tolist()
    ElasticInferenceAcceleratorsList = [x for x in ElasticInferenceAcceleratorsList if str(x) != 'nan']

    Ipv6AddressesList = df['Ipv6Addresses'].tolist()
    Ipv6AddressesList = [x for x in Ipv6AddressesList if str(x) != 'nan']

    LicenseSpecificationsList = df['LicenseSpecifications'].tolist()
    LicenseSpecificationsList = [x for x in LicenseSpecificationsList if str(x) != 'nan']

    NetworkInterfacesList = df['NetworkInterfaces'].tolist()
    NetworkInterfacesList = [x for x in NetworkInterfacesList if str(x) != 'nan']

    SecurityGroupIdsList = df['SecurityGroupIds'].tolist()
    SecurityGroupIdsList = [x for x in SecurityGroupIdsList if str(x) != 'nan']

    SecurityGroupsList = df['SecurityGroups'].tolist()
    SecurityGroupsList = [x for x in SecurityGroupsList if str(x) != 'nan']

    SsmAssociationsList = df['SsmAssociations'].tolist()
    SsmAssociationsList = [x for x in SsmAssociationsList if str(x) != 'nan']

    VolumesList = df['Volumes'].tolist()
    VolumesList = [x for x in VolumesList if str(x) != 'nan']

    Properties['BlockDeviceMappings'] = BDMapping
    Properties['ElasticGpuSpecifications'] = ElasticGpuSpecificationsList
    Properties['ElasticInferenceAccelerators'] = ElasticInferenceAcceleratorsList
    Properties['Ipv6Addresses'] = Ipv6AddressesList
    Properties['LicenseSpecifications'] = LicenseSpecificationsList
    Properties['NetworkInterfaces'] = NetworkInterfacesList
    Properties['SecurityGroupIds'] = SecurityGroupIdsList
    Properties['SecurityGroups'] = SecurityGroupsList
    Properties['SsmAssociations'] = SsmAssociationsList
    Properties['Tags'] = taglist
    Properties['Volumes'] = VolumesList

    filtered = {k: v for k, v in Properties.items() if v is not None and v != []}
    Properties.clear()
    Properties.update(filtered)

    ##print(json.dumps(Properties,indent=4))
    cfnDict["Resources"]["MyInstance"]["Properties"].update(Properties)
    cfnDict["Resources"]["MyInstance"]["Properties"].update(data)
    # Write JSON file
    with io.open('C:\\Users\\mayank\\Documents\\pythonWorkspace\\autogenerateCFN\\EC2Template.json', 'w',
                 encoding='utf8') as outfile:
        str_ = json.dumps(cfnDict,
                          indent=4, sort_keys=True,
                          separators=(',', ': '), ensure_ascii=False)
        outfile.write(str_)
    print("EC2 Create Function Executed...")
def create_sg():
    cfnDict = {
        "AWSTemplateFormatVersion": "2010-09-09",
        "Resources": {
            "MySecurityGroup": {
                "Type": "AWS::EC2::SecurityGroup",
                "Properties": {}
            }
        }
    }
    loc = ("C:\\Users\\mayank\\Documents\\pythonWorkspace\\autogenerateCFN\\EC2Details.xlsx")
    workbook = xlrd.open_workbook(loc, on_demand=True)
    worksheet = workbook.sheet_by_index(1)
    first_row = []  # The row where we stock the name of the column
    for col in range(worksheet.ncols):
        first_row.append(worksheet.cell_value(0, col))
    data = {}
    for row in range(1, 2):
        elm = {}
        for col in range(worksheet.ncols):
            if worksheet.cell_value(row, col) != '' and first_row[col] != 'EgressIpProtocol' and first_row[
                col] != 'EgressFromPort' \
                    and first_row[col] != 'EgressToPort' \
                    and first_row[col] != 'EgressCidrIp' and first_row[col] != 'IngressIpProtocol' and first_row[
                col] != 'IngressFromPort' \
                    and first_row[col] != 'IngressToPort' and first_row[col] != 'IngressCidrIp' and first_row[
                col] != 'TagsKey' \
                    and first_row[col] != 'TagsValue':
                elm[first_row[col]] = worksheet.cell_value(row, col)
        data = elm
    ##print (json.dumps(data,indent=4))
    Properties = {
        "SecurityGroupEgress": [],
        "SecurityGroupIngress": [],
        "Tags": []
    }
    Egress = []
    for i in range(1, worksheet.nrows):
        if worksheet.cell_value(i, 2) != '' and worksheet.cell_value(i, 3) != '' and worksheet.cell_value(i, 4) != '' \
                and worksheet.cell_value(i, 5) != '':
            ep = worksheet.cell_value(i, 2)
            efp = worksheet.cell_value(i, 3)
            etp = worksheet.cell_value(i, 4)
            ecid = worksheet.cell_value(i, 5)
            SecurityGroupEgressDict = {
                "IpProtocol": ep,
                "FromPort": str(int(efp)),
                "ToPort": str(int(etp)),
                "CidrIp": ecid
            }
            Egress.append(SecurityGroupEgressDict)
    # print(Egress)
    Properties["SecurityGroupEgress"] = Egress
    Ingress = []
    for i in range(1, worksheet.nrows):
        if worksheet.cell_value(i, 6) != '' and worksheet.cell_value(i, 7) != '' and worksheet.cell_value(i, 8) != '' \
                and worksheet.cell_value(i, 9) != '':
            inp = worksheet.cell_value(i, 6)
            infp = worksheet.cell_value(i, 7)
            intp = worksheet.cell_value(i, 8)
            incid = worksheet.cell_value(i, 9)
            SecurityGroupIngressDict = {
                "IpProtocol": inp,
                "FromPort": str(int(infp)),
                "ToPort": str(int(intp)),
                "CidrIp": incid
            }
            Ingress.append(SecurityGroupIngressDict)
    Properties["SecurityGroupIngress"] = Ingress

    taglist = []
    for j in range(2, worksheet.nrows):
        if worksheet.cell_value(j, 10) != '' and worksheet.cell_value(j, 11) != '':
            tdict = {}
            tdict['Key'] = worksheet.cell_value(j, 10)
            tdict['Value'] = worksheet.cell_value(j, 11)
            # print(tdict)
            taglist.append(tdict)
    ##print(taglist)
    Properties["Tags"] = taglist
    filtered = {k: v for k, v in Properties.items() if v is not None and v != []}
    Properties.clear()
    Properties.update(filtered)

    cfnDict["Resources"]["MySecurityGroup"]["Properties"].update(data)
    cfnDict["Resources"]["MySecurityGroup"]["Properties"].update(Properties)
    ##print(json.dumps(cfnDict,indent=4))

    # Write JSON file
    with io.open('C:\\Users\\mayank\\Documents\\pythonWorkspace\\autogenerateCFN\\SecurityGroupTemplate.json', 'w',
                 encoding='utf8') as outfile:
        str_ = json.dumps(cfnDict,
                          indent=4, sort_keys=True,
                          separators=(',', ': '), ensure_ascii=False)
        outfile.write(str_)
    print("SG Create Function Executed...")

if __name__ == "__main__":
    create_ec2()
    create_sg()
