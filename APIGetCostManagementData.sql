-- =============================================
-- [INV].[APIGetCostManagementData]
-- Returns 3 result sets:
--   #1 - Price List  (Sheet1)
--   #2 - Raw Cost    (Sheet2)
--   #3 - BOM         (Sheet3)
-- =============================================
create or alter procedure [INV].[APIGetCostManagementData]
    @Operation          nvarchar(50)   = 'GetAll',
    @User               nvarchar(100)  = null,
    @AppVersionWeb      nvarchar(20)   = null,
    @AppVersionAndroid  nvarchar(20)   = null,
    @AppVersionIos      nvarchar(20)   = null,
    @AppVersionDesktop  nvarchar(20)   = null,
    @PlatForm           nvarchar(20)   = null,
    @FireBaseToken      nvarchar(500)  = null,
    @SqlStatement       nvarchar(max)  = null,
    @State              int            = 1  out,
    @Message            nvarchar(500)  = ''  out
as
begin
    set nocount on;
    set @State   = 1;
    set @Message = '';

    if @Operation = 'GetAll'
    begin

        -- ── Result Set 1: Price List ──────────────────────────────────────────
        -- ItemCode, ItemDescription, PriceSellingUnit, SellingConversion
        select
             im.ItemCode
            ,im.ItemDescription
            ,isnull(pl.PriceSellingUnit, 0)     as PriceSellingUnit
            ,isnull(pl.SellingConversion, 1)    as SellingConversion
        from [INV].[ItemMaster] im
        inner join [INV].[PriceList] pl
            on pl.ItemID = im.ItemID
        where im.ItemType = 'F'          -- finished goods only
          and im.IsActive  = 1
        order by im.ItemCode;

        -- ── Result Set 2: Raw Cost ────────────────────────────────────────────
        -- ItemCode, ItemDescription, LastCost, AverageCost, ItemType
        select
             im.ItemCode
            ,im.ItemDescription
            ,isnull(im.LastCost, 0)      as LastCost
            ,isnull(im.AverageCost, 0)   as AverageCost
            ,im.ItemType
        from [INV].[ItemMaster] im
        where im.ItemType in ('R', 'P', 'S', 'B')  -- raw + semi-finished
          and im.IsActive = 1
        order by im.ItemType, im.ItemCode;

        -- ── Result Set 3: BOM ─────────────────────────────────────────────────
        -- ParentItemCode, ChildItemCode, Quantity, ParentBatchQty
        select
             parent.ItemCode    as ParentItemCode
            ,child.ItemCode     as ChildItemCode
            ,bl.Quantity
            ,bh.BatchQty        as ParentBatchQty
        from [INV].[BOMHeader] bh
        inner join [INV].[BOMLine]   bl     on bl.HeaderID  = bh.HeaderID
        inner join [INV].[ItemMaster] parent on parent.ItemID = bh.ParentItemID
        inner join [INV].[ItemMaster] child  on child.ItemID  = bl.ChildItemID
        where bh.IsActive = 1
          and bl.IsActive = 1
        order by parent.ItemCode, bl.LineNo;

    end
    else
    begin
        set @State   = 0;
        set @Message = 'Operation not found: ' + isnull(@Operation, '');
    end

end
