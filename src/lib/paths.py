
import os

SAP_SIGMANEST_PRD = r"\\hiifileserv1\sigmanestprd"
SAP_DATA_FILES = r"\\hssieng\SNData\SimTrans\SAP Data Files"
SIGMANEST_WORKORDERS = r"\\hssieng\DATA\HS\SAP - Material Master_BOM\SigmaNest Work Orders"

def workorder_file(job, shipment, extra_id=None):
    extra = ''
    if extra_id:
        extra = '_{}'.format(extra_id)

    return os.path.join(
        SIGMANEST_WORKORDERS,
        "20{} WOrk Orders Created".format(job[1:3]),
        "{}-{}_SimTrans_WO{}.xls".format(job, shipment, extra)
    )
