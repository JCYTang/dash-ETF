import win32com.client
from win32com.client import VARIANT
import pythoncom


class Iress:

    RequestState_DataReady = 2
    RequestState_DataError = 3
    PagingState_NoMoreData = 2
    UpdateState_UpdateReady = 2
    UpdateState_UpdateError = 3

    def __init__(self, method, fields=None, inputs=None):
        self.method = method
        self.fields = VARIANT(pythoncom.VT_VARIANT, fields)
        self.inputs = inputs
        self.requestManager = win32com.client.gencache.EnsureDispatch("IressServerApi.RequestManager", pythoncom.CoInitialize())
        self.requester = self.requestManager.CreateMethod("IRESS", "", method, 1)
        self.data = None

    def set_inputs(self):
        '''input_dict is a dictionary with input field names as the key and a list of input values as the key'''
        self.requester.Input.Header.Set("WaitForResponse", False)
        for field, values in self.inputs.items():
            input_field = VARIANT(pythoncom.VT_VARIANT, [field])
            input_values = VARIANT(pythoncom.VT_VARIANT, [[v] for v in values])
            self.requester.Input.Parameters.Set(input_field, input_values)

    def execute(self):
        self.requester.Execute()
        while True:
            if self.requester.RequestState == Iress.RequestState_DataReady:
                while True:
                    if self.requester.PagingState == Iress.PagingState_NoMoreData:
                        # All historical pages have been received.  The UpdateState is set to UpdateState_WaitingForUpdates until
                        # an update packet is received, upon which it will be set to UpdateState_UpdateReady.
                        break

                    # Get the next page
                    self.requester.Execute()

                break

        self.data = [list(d) for d in self.requester.Output.DataRows.GetRows(self.fields)]

    def retrieve_data(self):
        # retrieve latest data point
        if self.requester.UpdateState == Iress.UpdateState_UpdateReady:
            # Some update data has been received and can be retrieved.
            vAllUpdateData = self.requester.Output.UpdateRows.GetRowsAndRemove()
            if len(vAllUpdateData) > 0:
                update_data = self.requester.Output.UpdateRows.GetRowsFromRetrievedData(self.fields, vAllUpdateData)
                for u in update_data[0]:
                    for idx, d in enumerate(self.data):
                        if d[0] == u[0]:
                            self.data[idx] = list(u)

        return self.data


if __name__ == '__main__':
    # codes = ['MGOC', 'MGOCAUDINAV', 'MGOC-AUINAV']
    # exchanges = ['AXW', 'ETF', 'NGIF']
    codes = ['LSGE', 'LSGEAUDINAV', 'SPFUT']
    exchanges = ['AXW', 'ETF', 'ID']
    input_dict = {
        'SecurityCode': codes,
        'Exchange': exchanges,
    }
    method = 'pricingquoteexget'
    fields = ['SecurityCode', 'BidPrice', 'AskPrice', 'LastPrice', 'MovementPercent', 'BidVolume', 'AskVolume']
    iressobj = Iress(method, fields, input_dict)
    iressobj.set_inputs()
    iressobj.execute()
    iressobj.retrieve_data()