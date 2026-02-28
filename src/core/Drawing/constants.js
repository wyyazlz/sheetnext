// Drawing 模块公共常量

// 连接线类型（Excel标准）
// 箭头样式通过 shapeStyle.startArrow/endArrow 区分，不是通过不同的shapeType
export const CONNECTOR_TYPES = [
    'line',              // 直线
    'straightConnector1', // 直接连接符
    'bentConnector3',    // 肘形连接符
    'curvedConnector3'   // 曲线连接符
];
