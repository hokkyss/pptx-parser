Feature: Element ID Scoping, Uniqueness & Map Indexing
  As a developer building complex multi-slide presentations with @hokkyss/pptx
  I want element IDs to be indexed in O(1) time and validated strictly per slide
  So that I catch accidental ID collisions immediately while maintaining ECMA-376 OpenXML schema compliance

  Background:
    Given a new presentation instance is initialized
    And a slide is added to the presentation

  # =========================================================================
  # 1. MAP INDEXING & O(1) QUERYING
  # =========================================================================

  Scenario: Querying an element by ID in O(1) time
    Given a shape is added with custom id "hero-banner"
    When I query the slide with getElementById("hero-banner")
    Then the retrieved element should have id "hero-banner"
    And the element is returned in O(1) lookup time from the slide Map

  Scenario: Querying an element nested inside a group by ID
    Given a group container contains a child shape with id "nested-item"
    When I query the slide with getElementById("nested-item")
    Then the retrieved element should have id "nested-item"

  Scenario: Querying a non-existent element returns undefined
    When I query the slide with getElementById("does-not-exist")
    Then the result should be undefined

  # =========================================================================
  # 2. DUPLICATE ID DETECTION & EARLY BAILING
  # =========================================================================

  Scenario: Throwing early on duplicate element ID within the same slide
    Given a shape is added with id "metric-card"
    When I attempt to add another shape with id "metric-card" to the same slide
    Then an error should be thrown matching 'Duplicate element ID "metric-card" detected on Slide 1'

  Scenario Outline: Preventing duplicate IDs across mixed element types
    Given a <first_type> is added with id "shared-id"
    When I attempt to add a <second_type> with id "shared-id" to the same slide
    Then an error should be thrown matching 'Duplicate element ID "shared-id" detected on Slide 1'

    Examples:
      | first_type | second_type |
      | text box   | table       |
      | table      | chart       |
      | chart      | connector   |
      | connector  | image       |
      | image      | group       |

  Scenario: Throwing early when a group child collides with a top-level shape ID
    Given a top-level shape is added with id "card-alpha"
    When I attempt to add a group containing a child shape with id "card-alpha"
    Then an error should be thrown matching 'Duplicate element ID "card-alpha" detected on Slide 1'

  # =========================================================================
  # 3. ID LIFECYCLE & AUTO-GENERATION
  # =========================================================================

  Scenario: Reusing an element ID after removal
    Given a shape is added with id "temporary-box"
    When I call removeElement("temporary-box")
    Then removeElement returns true
    And getElementById("temporary-box") returns undefined
    And I can add a new shape with id "temporary-box" without throwing an error

  Scenario: Auto-increment counter skips manually claimed numeric IDs
    Given shapes are manually added with id "1" and id "2"
    When I add a shape without specifying an ID
    Then the auto-generated ID assigned to the shape is "3"
    And no ID collision error is thrown

  # =========================================================================
  # 4. CROSS-SLIDE ISOLATION & ECMA-376 SCHEMA MAPPING
  # =========================================================================

  Scenario: Permitting identical element IDs across different slides
    Given Slide 1 contains a shape with id "stage-1"
    When I add a shape with id "stage-1" to Slide 2
    Then both shapes are registered successfully without collision
    And querying getElementById("stage-1") on Slide 1 returns the Slide 1 element
    And querying getElementById("stage-1") on Slide 2 returns the Slide 2 element

  Scenario: Serializing developer string IDs to valid OpenXML unsigned integers
    Given a slide contains shapes with custom string IDs "card-alpha" and "card-beta"
    When the presentation is serialized to PPTX
    Then each shape's <p:cNvPr id="..."> in slide XML is an unsigned integer (e.g. 2, 3, 4)
    And the shape name attribute preserves the developer name "card-alpha"
    And the resulting presentation passes strict OpenXML schema validation without repair warnings
