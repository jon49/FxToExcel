namespace FXCore

module Structure =
    
    ///Transaction 
    type public Transaction = {Id:string; Commodity:string; Date:System.DateTime; Account:string; Amount:float; Payee:string
                                ; Category:string; Number:string; Memo:string; Shares:float; Category_Type: string}
    
    ///List of transactions
    type public Transactions = {Transactions: Transaction list}
        



        
    