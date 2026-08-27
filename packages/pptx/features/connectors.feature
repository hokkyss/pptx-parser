Feature: Shape Connectors & Arrowhead Customization
  As a presentation developer using @hokkyss/pptx
  I want to connect shapes with customizable arrows and anchor lines to specific shape edges
  So that I can build interactive diagrams, flowcharts, and architecture pipelines in PowerPoint

  Background:
    Given a new presentation instance is initialized
    And a slide is added to the presentation

  # =========================================================================
  # 1. SHAPE ATTACHMENT & COMPASS ANCHORING
  # =========================================================================

  Scenario: Attaching a connector between two shapes using compass positions
    Given a shape exists with id "card-1" at x=1", y=1", w=2", h=2"
    And a shape exists with id "card-2" at x=5", y=1", w=2", h=2"
    When I add a connector with:
      | from_shapeId  | card-1 |
      | from_position | right  |
      | to_shapeId    | card-2 |
      | to_position   | left   |
    Then the connector start anchor is bound to shape "card-1" at position "right"
    And the connector end anchor is bound to shape "card-2" at position "left"
    And the computed start coordinate is (x: 3", y: 2")
    And the computed end coordinate is (x: 5", y: 2")
    And the connector dimension cx is 2" and cy is 0"

  Scenario: Attaching a vertical connector between top and bottom anchors
    Given a shape exists with id "top-box" at x=2", y=1", w=4", h=2"
    And a shape exists with id "bottom-box" at x=2", y=5", w=4", h=2"
    When I add a connector from "top-box" (bottom) to "bottom-box" (top)
    Then the computed start coordinate is (x: 4", y: 3")
    And the computed end coordinate is (x: 4", y: 5")
    And the connector dimension cx is 0" and cy is 2"

  Scenario: Attaching a connector to shapes nested inside composite groups
    Given a group container "group-1" containing a child shape with id "stage-1" at x=1", y=2", w=3", h=3"
    And a group container "group-2" containing a child shape with id "stage-2" at x=5", y=2", w=3", h=3"
    When I add a connector from shapeId "stage-1" (right) to shapeId "stage-2" (left)
    Then the connector recursively locates "stage-1" inside "group-1"
    And the connector recursively locates "stage-2" inside "group-2"
    And the connector connects the two child shapes without error

  Scenario: Attempting to attach a connector directly to a group container fails
    Given a group container exists with id "pipeline-group"
    When I attempt to add a connector from shapeId "pipeline-group" to coordinate (x: 5", y: 5")
    Then an error should be thrown matching 'Cannot attach connector to a group ("pipeline-group")'

  Scenario: Attempting to attach a connector to a non-existent shape ID fails
    When I attempt to add a connector from shapeId "missing-shape" to coordinate (x: 5", y: 5")
    Then an error should be thrown matching 'Shape with id "missing-shape" was not found on this slide'

  # =========================================================================
  # 2. ARROWHEAD MARKERS & LINE FORMATTING
  # =========================================================================

  Scenario Outline: Customizing arrowhead types, lengths, and widths
    Given a shape exists with id "src-node" at x=1", y=1", w=2", h=2"
    And a shape exists with id "dst-node" at x=5", y=1", w=2", h=2"
    When I add a connector from "src-node" to "dst-node" with:
      | color      | 0284C7       |
      | dashStyle  | <dash_style> |
      | startArrow | <start_type> |
      | endArrow   | <end_type>   |
      | width      | <width>      |
      | length     | <length>     |
    Then the connector line property contains headEnd type "<xml_head>" with w="<xml_w>" and len="<xml_len>"
    And the connector line property contains tailEnd type "<xml_tail>" with w="<xml_w>" and len="<xml_len>"

    Examples:
      | dash_style | start_type | end_type | width | length | xml_head | xml_tail | xml_w | xml_len |
      | solid      | oval       | triangle | lg    | lg     | triangle | oval     | lg    | lg      |
      | dash       | diamond    | stealth  | sm    | med    | stealth  | diamond  | sm    | med     |
      | sysDot     | open       | none     | med   | sm     | none     | open     | med   | sm      |

  Scenario: Specifying detailed headEnd and tailEnd objects
    Given a shape exists with id "node-a" at x=1", y=1", w=2", h=2"
    And a shape exists with id "node-b" at x=5", y=1", w=2", h=2"
    When I add a connector from "node-a" to "node-b" with headEnd "{ type: 'stealth', width: 'lg', length: 'lg' }" and tailEnd "{ type: 'oval', width: 'sm', length: 'sm' }"
    Then the connector headEnd has type="stealth", width="lg", and length="lg"
    And the connector tailEnd has type="oval", width="sm", and length="sm"

  # =========================================================================
  # 3. OPENXML SERIALIZATION & POWERPOINT INTERACTION
  # =========================================================================

  Scenario: Generating OpenXML DrawingML connection elements and locking tags
    Given a slide contains connected shapes "box-a" and "box-b"
    When the presentation is exported to PPTX
    Then the connector XML contains <a:stCxn id="..." idx="..."/>
    And the connector XML contains <a:endCxn id="..." idx="..."/>
    And the connector XML contains <a:cxnSpLocks/> inside <p:cNvCxnSpPr>
    And opening the presentation in PowerPoint allows live dragging of shapes with stretching connectors
